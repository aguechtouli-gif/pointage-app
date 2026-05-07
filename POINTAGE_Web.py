#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
web.py – Application web de pointage (Flask)
Partage la même base de données que l'application Tkinter.
Auteur : TRC
Date : 2024
"""

import os
import sqlite3
from datetime import datetime, date, timedelta
from functools import wraps
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, jsonify, send_file, make_response, send_from_directory
)
from werkzeug.utils import secure_filename
import pandas as pd
import hashlib
import tempfile


# Import de la classe Database depuis database.py
from database import Database
#from utils import paginate
from flask import session
import pdfkit
from time import time



PATH_WKHTMLTOPDF = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"   # adaptez si nécessaire
config = pdfkit.configuration(wkhtmltopdf=PATH_WKHTMLTOPDF)

import threading
import queue
import psutil
import os
from datetime import datetime, date

# Dictionnaire des tâches d'import en cours
import_tasks = {}  # task_id -> {'status': 'pending', 'progress': 0, 'total': 0, 'message': '', 'queue': queue.Queue()}

def get_system_stats():
    """Retourne les statistiques système (CPU, RAM)"""
    try:
        return {
            'cpu_percent': psutil.cpu_percent(interval=0.5),
            'ram_percent': psutil.virtual_memory().percent,
            'ram_used': psutil.virtual_memory().used // (1024**2),  # MB
            'ram_total': psutil.virtual_memory().total // (1024**2)  # MB
        }
    except:
        return {
            'cpu_percent': 0,
            'ram_percent': 0,
            'ram_used': 0,
            'ram_total': 0
        }

def import_task_worker(task_id, filepath, format_colonnes, user_id):
    """Tâche d'import exécutée en arrière-plan"""
    try:
        import_tasks[task_id]['status'] = 'running'
        import_tasks[task_id]['message'] = 'Lecture du fichier...'
        import_tasks[task_id]['queue'].put({'type': 'progress', 'value': 10})
        
        results = db.import_pointages_from_file(
            filepath=filepath,
            format_colonnes=format_colonnes,
            user_id=user_id,
            progress_callback=lambda p, msg: update_progress(task_id, p, msg)
        )
        
        import_tasks[task_id]['status'] = 'completed'
        import_tasks[task_id]['progress'] = 100
        import_tasks[task_id]['message'] = f"Import terminé : {results['importes']} pointages"
        import_tasks[task_id]['results'] = results
        import_tasks[task_id]['total'] = results.get('total', 0)
        import_tasks[task_id]['queue'].put({'type': 'complete', 'results': results})
        
    except Exception as e:
        import_tasks[task_id]['status'] = 'failed'
        import_tasks[task_id]['message'] = str(e)
        import_tasks[task_id]['queue'].put({'type': 'error', 'message': str(e)})

def update_progress(task_id, progress, message):
    """Met à jour la progression d'une tâche"""
    if task_id in import_tasks:
        import_tasks[task_id]['progress'] = progress
        import_tasks[task_id]['message'] = message
        import_tasks[task_id]['queue'].put({'type': 'progress', 'value': progress, 'message': message})

#-----------------------------------------------------------------
# Fonction utilitaire pour convertir un DataFrame en PDF
#-----------------------------------------------------------------
def dataframe_to_pdf(df, titre, filename, orientation='portrait'):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm

    pagesize = landscape(A4) if orientation == 'landscape' else A4
    doc = SimpleDocTemplate(filename, pagesize=pagesize,
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    elements = []
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    elements.append(Paragraph(titre, title_style))
    elements.append(Spacer(1, 0.5*cm))
    data = [df.columns.tolist()] + df.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BACKGROUND', (0,1), (-1,-1), colors.beige),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
    ]))
    elements.append(table)
    doc.build(elements)
# ============================================================
# Fonction utilitaire pour convertir une heure en minutes
# ============================================================
def heure_en_minutes(heure_str):
    try:
        parts = heure_str.split(':')
        h = int(parts[0])
        m = int(parts[1])
        s = int(parts[2]) if len(parts) > 2 else 0
        return h * 60 + m + s / 60
    except:
        return 0

app = Flask(__name__)
app.secret_key = "une_cle_secrete_tres_longue_et_unique_changez_moi"
app.config['UPLOAD_FOLDER'] = os.path.join('pointage_data', 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db = Database()

# -------------------------------------------------------------------
# Décorateurs
# -------------------------------------------------------------------
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash("Veuillez vous connecter pour accéder à cette page.", "warning")
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get('role') != 'admin':
            flash("Accès réservé aux administrateurs.", "danger")
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

# -------------------------------------------------------------------
# Authentification
# -------------------------------------------------------------------
@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = db.authenticate_user(username, password)
        if user:
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['nom'] = user['nom']
            session['prenom'] = user['prenom']
            session['role'] = user['role']
            flash(f"Bienvenue {user['prenom']} {user['nom']} !", "success")
            return redirect(url_for('dashboard'))
        else:
            flash("Nom d'utilisateur ou mot de passe incorrect.", "danger")
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash("Vous avez été déconnecté.", "info")
    return redirect(url_for('login'))

# -------------------------------------------------------------------
# Tableau de bord
# -------------------------------------------------------------------
@app.route('/dashboard')
@login_required
def dashboard():
    stats = db.get_statistiques_completes()
    derniers_pointages = stats.get('derniers_pointages', [])
    with db.get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM personnel WHERE statut = 'actif'")
        nb_personnel_actif = c.fetchone()[0]
    return render_template('dashboard.html', 
                           stats=stats, 
                           pointages=derniers_pointages,
                           nb_personnel_actif=nb_personnel_actif)

# -------------------------------------------------------------------
# Gestion du personnel
# -------------------------------------------------------------------
@app.route('/personnel')
@login_required
def liste_personnel():
    # Pagination
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    search = request.args.get('search', '')
    
    # Récupération des données
    if search:
        personnel = db.search_personnel(search)
    else:
        personnel = db.get_personnel()
    
    # Pagination
    total = len(personnel)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    personnel_page = personnel[offset:offset+per_page]
    
    # Filtres pour la pagination
    filter_args = {'search': search} if search else {}
    
    # Charger les listes pour les filtres (optionnel)
    directions = db.get_all_directions()
    departements = db.get_all_departements()
    services = db.get_all_services()
    equipes = db.get_all_equipes()
    
    return render_template('personnel/liste.html',
                           personnel=personnel_page,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args,
                           search=search,
                           directions=directions,
                           departements=departements,
                           services=services,
                           equipes=equipes)

@app.route('/personnel/<int:pid>')
@login_required
def detail_personnel(pid):
    pers = db.get_personnel(pid)
    if not pers:
        flash("Personnel non trouvé.", "danger")
        return redirect(url_for('liste_personnel'))
    pointages = db.get_pointages_filtres(matricule=pers['matricule'])
    for pt in pointages:
        if pt.get('type_pointage') == 'entrée':
            minutes_pt = heure_en_minutes(pt['heure_pointage'])
            minutes_th = heure_en_minutes(pers['heure_entree_theorique'])
            tolerance = pers.get('tolerance_entree', 0)
            retard_brut = minutes_pt - minutes_th
            if retard_brut > tolerance:
                pt['minutes_retard'] = retard_brut
            else:
                pt['minutes_retard'] = 0
        else:
            pt['minutes_retard'] = 0
    return render_template('personnel/details.html', pers=pers, pointages=pointages[:20])


@app.route('/personnel/ajouter', methods=['GET', 'POST'])
@login_required
def ajouter_personnel():
    if request.method == 'POST':
        try:
            matricule = request.form.get('matricule', '').strip()
            if not matricule:
                flash("Le matricule est obligatoire.", "danger")
                return redirect(url_for('ajouter_personnel'))
            badge_id = request.form.get('badge_id') or None
            nom = request.form.get('nom', '').upper().strip()
            prenom = request.form.get('prenom', '').title().strip()
            type_person = request.form.get('type_person', 'cadre')
            fonction = request.form.get('fonction', '').strip() or None
            activite_id = request.form.get('activite_id') or None
            division_id = request.form.get('division_id') or None
            direction_id = request.form.get('direction_id') or None
            departement_id = request.form.get('departement_id') or None
            service_id = request.form.get('service_id') or None
            equipe_id = request.form.get('equipe_id') or None
            quart_id = request.form.get('quart_id') or None
            date_embauche = request.form.get('date_embauche') or None
            date_naissance = request.form.get('date_naissance') or None
            adresse = request.form.get('adresse') or None
            telephone = request.form.get('telephone') or None
            email = request.form.get('email') or None
            statut = request.form.get('statut', 'actif')
            concerne_pointage = 1 if 'concerne_pointage' in request.form else 0
            type_quart = request.form.get('type_quart', 'jour')
            heure_entree = request.form.get('heure_entree_theorique', '08:00:00')
            heure_sortie = request.form.get('heure_sortie_theorique', '17:00:00')
            tolerance_entree = int(request.form.get('tolerance_entree', 0))
            tolerance_sortie = int(request.form.get('tolerance_sortie', 0))
            photo = None
            if 'photo' in request.files:
                file = request.files['photo']
                if file and file.filename:
                    photo = db.sauvegarder_photo(file, matricule)
            data = (
                matricule, badge_id, nom, prenom, type_person, fonction,
                activite_id, division_id, direction_id, departement_id, service_id, equipe_id,
                date_embauche, date_naissance, adresse, telephone, email, photo, statut,
                concerne_pointage, type_quart, heure_entree, heure_sortie,
                tolerance_entree, tolerance_sortie
            )
            db.add_personnel(data)
            flash("Personnel ajouté avec succès.", "success")
            return redirect(url_for('liste_personnel'))
        except Exception as e:
            flash(f"Erreur lors de l'ajout : {str(e)}", "danger")
            return redirect(url_for('ajouter_personnel'))
    activites = db.get_all_activites()
    divisions = db.get_all_divisions()
    directions = db.get_all_directions()
    departements = db.get_all_departements()
    services = db.get_all_services()
    equipes = db.get_all_equipes()
    quarts = db.get_all_quarts()
    types_personnel = db.get_all_types_personnel()
    return render_template('personnel/ajouter.html',
                           activites=activites,
                           divisions=divisions,
                           directions=directions,
                           departements=departements,
                           services=services,
                           equipes=equipes,
                           quarts=quarts,
                           types_personnel=types_personnel)

@app.route('/personnel/<int:pid>/modifier', methods=['GET', 'POST'])
@login_required
def modifier_personnel(pid):
    pers = db.get_personnel(pid)
    if not pers:
        flash("Personnel non trouvé.", "danger")
        return redirect(url_for('liste_personnel'))
    if request.method == 'POST':
        try:
            matricule = request.form.get('matricule', '').strip()
            if not matricule:
                flash("Le matricule est obligatoire.", "danger")
                return redirect(url_for('modifier_personnel', pid=pid))
            badge_id = request.form.get('badge_id') or None
            nom = request.form.get('nom', '').upper().strip()
            prenom = request.form.get('prenom', '').title().strip()
            type_person = request.form.get('type_person', 'cadre')
            fonction = request.form.get('fonction', '').strip() or None
            activite_id = request.form.get('activite_id') or None
            division_id = request.form.get('division_id') or None
            direction_id = request.form.get('direction_id') or None
            departement_id = request.form.get('departement_id') or None
            service_id = request.form.get('service_id') or None
            equipe_id = request.form.get('equipe_id') or None
            quart_id = request.form.get('quart_id') or None
            date_embauche = request.form.get('date_embauche') or None
            date_naissance = request.form.get('date_naissance') or None
            adresse = request.form.get('adresse') or None
            telephone = request.form.get('telephone') or None
            email = request.form.get('email') or None
            statut = request.form.get('statut', 'actif')
            concerne_pointage = 1 if 'concerne_pointage' in request.form else 0
            type_quart = request.form.get('type_quart', 'jour')
            heure_entree = request.form.get('heure_entree_theorique', '08:00:00')
            heure_sortie = request.form.get('heure_sortie_theorique', '17:00:00')
            tolerance_entree = int(request.form.get('tolerance_entree', 0))
            tolerance_sortie = int(request.form.get('tolerance_sortie', 0))
            photo = pers['photo']
            if 'photo' in request.files:
                file = request.files['photo']
                if file and file.filename:
                    photo = db.sauvegarder_photo(file, matricule)
            data = (
                matricule, badge_id, nom, prenom, type_person, fonction,
                activite_id, division_id, direction_id, departement_id, service_id, equipe_id,
                date_embauche, date_naissance, adresse, telephone, email, photo, statut,
                concerne_pointage, type_quart, heure_entree, heure_sortie,
                tolerance_entree, tolerance_sortie
            )
            db.update_personnel(pid, data)
            flash("Modifications enregistrées.", "success")
            return redirect(url_for('detail_personnel', pid=pid))
        except Exception as e:
            flash(f"Erreur : {e}", "danger")
            return redirect(url_for('modifier_personnel', pid=pid))
    activites = db.get_all_activites()
    divisions = db.get_all_divisions()
    directions = db.get_all_directions()
    departements = db.get_all_departements()
    services = db.get_all_services()
    equipes = db.get_all_equipes()
    quarts = db.get_all_quarts()
    types_personnel = db.get_all_types_personnel()
    return render_template('personnel/modifier.html',
                           pers=pers,
                           activites=activites,
                           divisions=divisions,
                           directions=directions,
                           departements=departements,
                           services=services,
                           equipes=equipes,
                           quarts=quarts,
                           types_personnel=types_personnel)

@app.route('/personnel/<int:pid>/supprimer', methods=['POST'])
@login_required
@admin_required
def supprimer_personnel(pid):
    if db.delete_personnel(pid):
        flash("Personnel supprimé.", "success")
    else:
        flash("Erreur lors de la suppression.", "danger")
    return redirect(url_for('liste_personnel'))

@app.route('/personnel/importer', methods=['GET', 'POST'])
@login_required
@admin_required
def importer_personnel():
    if request.method == 'POST':
        fichier = request.files['fichier']
        if fichier and fichier.filename.endswith(('.csv', '.xlsx', '.xls')):
            filename = secure_filename(fichier.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            fichier.save(filepath)
            resultats = db.import_personnel_from_file(filepath)
            flash(f"Import terminé : {resultats['importes']} importés, {resultats['mis_a_jour']} mis à jour, {resultats['erreurs']} erreurs.", "info")
            return redirect(url_for('liste_personnel'))
        else:
            flash("Format de fichier non supporté.", "danger")
    return render_template('personnel/import.html')

# ------------------------------------------------------------
# Routes pour filtrer le personnel par hiérarchie
# ------------------------------------------------------------

#-------------------------------------------------------------
# Personnel par Activité
#--------------------------------------------------------------
@app.route('/personnel/activite/<int:act_id>')
@login_required
def personnel_par_activite(act_id):
    # Pagination
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    # Récupérer toutes les données
    personnel = db.get_personnel_par_activite(act_id)
    nom = db.get_activite(act_id)['nom_activite'] if db.get_activite(act_id) else 'Activité'
    
    # Pagination
    total = len(personnel)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    personnel_page = personnel[offset:offset+per_page]
    
    filter_args = {'act_id': act_id}
    
    return render_template('personnel/liste.html', 
                           personnel=personnel_page,
                           search=nom,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args)

#------------------------------------------------------------
# Personnel par Division
#------------------------------------------------------------
@app.route('/personnel/division/<int:div_id>')
@login_required
def personnel_par_division(div_id):
    # Pagination
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    personnel = db.get_personnel_par_division(div_id)
    nom = db.get_division(div_id)['nom_division'] if db.get_division(div_id) else 'Division'
    
    total = len(personnel)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    personnel_page = personnel[offset:offset+per_page]
    
    filter_args = {'div_id': div_id}
    
    return render_template('personnel/liste.html',
                           personnel=personnel_page,
                           search=nom,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args)
    
#------------------------------------------------------------
# Personnel par Direction
#------------------------------------------------------------
@app.route('/personnel/direction/<int:direction_id>')
@login_required
def personnel_par_direction(direction_id):
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    personnel = db.get_personnel_par_direction(direction_id)
    nom = db.get_direction(direction_id)['nom_direction'] if db.get_direction(direction_id) else 'Direction'
    
    total = len(personnel)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    personnel_page = personnel[offset:offset+per_page]
    
    filter_args = {'direction_id': direction_id}
    
    return render_template('personnel/liste.html',
                           personnel=personnel_page,
                           search=nom,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args)
    
#------------------------------------------------------------
# Personnel par Département
#------------------------------------------------------------
@app.route('/personnel/departement/<int:departement_id>')
@login_required
def personnel_par_departement(departement_id):
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    personnel = db.get_personnel_par_departement(departement_id)
    nom = db.get_departement(departement_id)['nom_departement'] if db.get_departement(departement_id) else 'Département'
    
    total = len(personnel)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    personnel_page = personnel[offset:offset+per_page]
    
    filter_args = {'departement_id': departement_id}
    
    return render_template('personnel/liste.html',
                           personnel=personnel_page,
                           search=nom,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args)

#------------------------------------------------------------
# Personnel par Service
#------------------------------------------------------------
@app.route('/personnel/service/<int:service_id>')
@login_required
def personnel_par_service(service_id):
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    personnel = db.get_personnel_par_service(service_id)
    nom = db.get_service(service_id)['nom_service'] if db.get_service(service_id) else 'Service'
    
    total = len(personnel)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    personnel_page = personnel[offset:offset+per_page]
    
    filter_args = {'service_id': service_id}
    
    return render_template('personnel/liste.html',
                           personnel=personnel_page,
                           search=nom,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args)

#------------------------------------------------------------
# Personnel par Équipe
#------------------------------------------------------------
@app.route('/personnel/equipe/<int:equipe_id>')
@login_required
def personnel_par_equipe(equipe_id):
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    personnel = db.get_personnel_par_equipe(equipe_id)
    equipe = db.get_equipe(equipe_id)
    if equipe:
        nom = equipe.get('nom_equipe', 'Équipe')
    else:
        nom = 'Équipe'
    
    
    total = len(personnel)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    personnel_page = personnel[offset:offset+per_page]
    
    filter_args = {'equipe_id': equipe_id}
    
    return render_template('personnel/liste.html',
                           personnel=personnel_page,
                           search=nom,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args)

#------------------------------------------------------------
# Personnel par Quart de travail
#------------------------------------------------------------
@app.route('/personnel/quart/<int:quart_id>')
@login_required
def personnel_par_quart(quart_id):
    with db.get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT nom_quart FROM quarts_travail WHERE id = ?", (quart_id,))
        row = c.fetchone()
        nom = row[0] if row else f"Quart {quart_id}"
    with db.get_connection() as conn:
        c = conn.cursor()
        c.execute("""
            SELECT p.*,
                   a.nom_activite as activite_nom,
                   d.nom_division as division_nom,
                   dir.nom_direction as direction_nom,
                   dep.nom_departement as departement_nom,
                   s.nom_service as service_nom,
                   e.nom_equipe as equipe_nom,
                   q.nom_quart as quart_nom
            FROM personnel p
            LEFT JOIN activites a ON p.activite_id = a.id
            LEFT JOIN divisions d ON p.division_id = d.id
            LEFT JOIN directions dir ON p.direction_id = dir.id
            LEFT JOIN departements dep ON p.departement_id = dep.id
            LEFT JOIN services s ON p.service_id = s.id
            LEFT JOIN equipes e ON p.equipe_id = e.id
            LEFT JOIN quarts_travail q ON p.quart_id = q.id
            WHERE p.quart_id = ? AND p.statut = 'actif'
            ORDER BY p.nom, p.prenom
        """, (quart_id,))
        personnel = [dict(row) for row in c.fetchall()]
    return render_template('personnel/liste.html', personnel=personnel, search=nom)

# ------------------------------------------------------------
# Export PDF d'un agent
# ------------------------------------------------------------
@app.route('/personnel/<int:pid>/export_pdf')
@login_required
def export_personnel_pdf(pid):
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    import tempfile, os

    pers = db.get_personnel(pid)
    if not pers:
        flash("Personnel non trouvé.", "danger")
        return redirect(url_for('liste_personnel'))

    pointages = db.get_pointages_filtres(matricule=pers['matricule'])[:20]

    titre = f"Rapport de {pers['prenom']} {pers['nom']} - {pers['fonction'] or ''} - {pers.get('direction_nom', '')}"

    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        pdf_path = tmp.name

    doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle('CustomTitle', parent=styles['Title'],
                                 fontSize=16, textColor=colors.darkblue,
                                 alignment=1, spaceAfter=20)
    style_normal = styles['Normal']

    elements = []

    logo_path = os.path.join('static', 'images', 'logo.png')
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=2*cm, height=2*cm)
        elements.append(logo)

    elements.append(Paragraph(titre, style_title))
    elements.append(Spacer(1, 0.5*cm))

    infos = [
        ["Matricule", pers['matricule']],
        ["Date d'embauche", pers['date_embauche'] or 'Non renseignée'],
        ["Téléphone", pers['telephone'] or ''],
        ["Email", pers['email'] or '']
    ]
    table_infos = Table(infos, colWidths=[4*cm, 8*cm])
    table_infos.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
    ]))
    elements.append(table_infos)
    elements.append(Spacer(1, 0.5*cm))

    if pointages:
        data = [["Date", "Heure", "Type"]]
        for pt in pointages:
            data.append([pt['date_pointage'], pt['heure_pointage'], pt['type_pointage']])
        table_pt = Table(data, repeatRows=1)
        table_pt.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('FONTSIZE', (0,0), (-1,-1), 9),
        ]))
        elements.append(table_pt)
    else:
        elements.append(Paragraph("Aucun pointage enregistré.", style_normal))

    from datetime import datetime
    user = session.get('prenom', '') + ' ' + session.get('nom', '')
    def add_footer(canvas, doc):
        canvas.saveState()
        footer = f"Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')} par {user}"
        canvas.setFont('Helvetica', 8)
        canvas.drawCentredString(doc.pagesize[0]/2, 1.5*cm, footer)
        canvas.restoreState()

    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
    return send_file(pdf_path, as_attachment=False, download_name=f"rapport_{pers['matricule']}.pdf")

# -------------------------------------------------------------------
# Gestion des pointages
# -------------------------------------------------------------------
@app.route('/pointages')
@login_required
def liste_pointages():
    # Pagination
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    # Récupération des filtres
    date_debut = request.args.get('date_debut', '')
    date_fin = request.args.get('date_fin', '')
    matricule = request.args.get('matricule', '')
    type_pointage = request.args.get('type_pointage', '')
    direction = request.args.get('direction', '')
    departement = request.args.get('departement', '')
    service = request.args.get('service', '')
    equipe = request.args.get('equipe', '')
    
    # Appel à la base de données (avec pagination directe pour les pointages)
    if date_debut and date_fin:
        pointages = db.get_pointages_filtres(
            date_debut=date_debut, date_fin=date_fin,
            matricule=matricule, type_pointage=type_pointage,
            direction_nom=direction, departement_nom=departement,
            service_nom=service, equipe_nom=equipe
        )
    else:
        pointages = []
    
    # Pagination en mémoire
    total = len(pointages)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    pointages_page = pointages[offset:offset+per_page]
    
    # Filtres pour la pagination
    filter_args = {
        'date_debut': date_debut,
        'date_fin': date_fin,
        'matricule': matricule,
        'type_pointage': type_pointage,
        'direction': direction,
        'departement': departement,
        'service': service,
        'equipe': equipe
    }
    filter_args = {k: v for k, v in filter_args.items() if v}
    
    # Charger les listes pour les filtres
    directions = db.get_all_directions()
    departements = db.get_all_departements()
    services = db.get_all_services()
    equipes = db.get_all_equipes()
    
    return render_template('pointages/liste.html',
                           pointages=pointages_page,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args,
                           date_debut=date_debut,
                           date_fin=date_fin,
                           matricule=matricule,
                           type_pointage=type_pointage,
                           direction=direction,
                           departement=departement,
                           service=service,
                           equipe=equipe,
                           directions=directions,
                           departements=departements,
                           services=services,
                           equipes=equipes)
#------------------------------------------------------------
# Pointage rapide (sans sélection de l'agent, juste matricule/badge et type de pointage)
#--------------------------------------------------------------------
@app.route('/pointages/rapide', methods=['GET', 'POST'])
@login_required
def pointage_rapide():
    if request.method == 'POST':
        identifiant = request.form['identifiant'].strip()
        type_pt = request.form['type_pointage']
        pers = db.get_personnel_by_matricule(identifiant) or db.get_personnel_by_badge(identifiant)
        if not pers:
            flash("Personnel non trouvé.", "danger")
        else:
            pid, msg = db.add_pointage_avance(
                matricule=pers['matricule'],
                type_pointage=type_pt,
                mode='web',
                user_id=session['user_id']
            )
            if pid:
                flash(f"Pointage enregistré pour {pers['prenom']} {pers['nom']} - {msg}", "success")
            else:
                flash(f"Erreur : {msg}", "danger")
        return redirect(url_for('pointage_rapide'))
    return render_template('pointages/rapide.html')

@app.route('/pointages/importer', methods=['GET', 'POST'])
@login_required
@admin_required
def importer_pointages():
    if request.method == 'POST':
        fichier = request.files['fichier']
        if fichier and fichier.filename.endswith(('.csv', '.xlsx', '.xls')):
            filename = secure_filename(fichier.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            fichier.save(filepath)
            mapping = {}
            resultats = db.import_pointages_from_file(filepath, format_colonnes=mapping, user_id=session['user_id'])
            flash(f"Import : {resultats['importes']} créés, {resultats['doublons']} doublons, {resultats['erreurs']} erreurs.", "info")
            if resultats['details']:
                for detail in resultats['details'][:20]:
                    flash(detail, "warning")
            return redirect(url_for('liste_pointages'))
        else:
            flash("Format de fichier non supporté. Utilisez CSV ou Excel.", "danger")
    return render_template('pointages/import.html')

# ------------------------------------------------------------
# API pour les menus hiérarchiques dépendants
# ------------------------------------------------------------
@app.route('/api/hierarchie/<type_entite>')
@login_required
def api_hierarchie(type_entite):
    """Retourne la liste des entités d'un type donné."""
    if type_entite == 'activites':
        data = db.get_all_activites()
    elif type_entite == 'divisions':
        data = db.get_all_divisions()
    elif type_entite == 'directions':
        data = db.get_all_directions()
    elif type_entite == 'departements':
        data = db.get_all_departements()
    elif type_entite == 'services':
        data = db.get_all_services()
    elif type_entite == 'equipes':
        data = db.get_all_equipes()
    else:
        return jsonify([])
    return jsonify(data)

@app.route('/api/hierarchie/enfants/<type_parent>/<int:parent_id>')
@login_required
def api_hierarchie_enfants(type_parent, parent_id):
    """Retourne les entités enfants (ex: divisions d'une activité, directions d'une division, etc.)"""
    if type_parent == 'activite':
        enfants = [d for d in db.get_all_divisions() if d.get('activite_id') == parent_id]
    elif type_parent == 'division':
        enfants = [d for d in db.get_all_directions() if d.get('division_id') == parent_id]
    elif type_parent == 'direction':
        enfants = [d for d in db.get_all_departements() if d.get('direction_id') == parent_id]
    elif type_parent == 'departement':
        enfants = [d for d in db.get_all_services() if d.get('departement_id') == parent_id]
    elif type_parent == 'service':
        enfants = [d for d in db.get_all_equipes() if d.get('service_id') == parent_id]
    else:
        enfants = []
    return jsonify(enfants)

# ------------------------------------------------------------
# Affichage du personnel par hiérarchie
# ------------------------------------------------------------
@app.route('/personnel/hierarchie')
@login_required
def personnel_hierarchie():
    """Page de sélection hiérarchique et affichage du personnel filtré."""
    entite_type = request.args.get('type')        # activite, division, direction, departement, service, equipe
    entite_id = request.args.get('id', type=int)
    inclure_descendants = request.args.get('descendants', 'false') == 'true'
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)

    if not entite_type or not entite_id:
        # Afficher la sélection initiale
        activites = db.get_all_activites()
        divisions = db.get_all_divisions()
        directions = db.get_all_directions()
        departements = db.get_all_departements()
        services = db.get_all_services()
        equipes = db.get_all_equipes()
        return render_template('personnel/hierarchie.html', 
                               activites=activites, 
                               divisions=divisions, 
                               directions=directions, 
                               departements=departements, 
                               services=services, 
                               equipes=equipes, 
                               personnel=None,)

    # Récupérer le nom de l'entité
    nom_entite = ''
    if entite_type == 'activite':
        ent = db.get_activite(entite_id)
        nom_entite = ent.get('nom_activite') if ent else ''
    elif entite_type == 'division':
        ent = db.get_division(entite_id)
        nom_entite = ent.get('nom_division') if ent else ''
    elif entite_type == 'direction':
        ent = db.get_direction(entite_id)
        nom_entite = ent.get('nom_direction') if ent else ''
    elif entite_type == 'departement':
        ent = db.get_departement(entite_id)
        nom_entite = ent.get('nom_departement') if ent else ''
    elif entite_type == 'service':
        ent = db.get_service(entite_id)
        nom_entite = ent.get('nom_service') if ent else ''
    elif entite_type == 'equipe':
        ent = db.get_equipe(entite_id)
        nom_entite = ent.get('nom_equipe') if ent else ''

    # Récupérer les IDs des entités à inclure (descendants si demandé)
    ids_inclus = [entite_id]
    if inclure_descendants:
        if entite_type == 'activite':
            # Toutes les divisions de cette activité, puis directions de ces divisions, etc.
            divisions = [d for d in db.get_all_divisions() if d.get('activite_id') == entite_id]
            ids_inclus.extend([d['id'] for d in divisions])
            directions = [d for d in db.get_all_directions() if d.get('division_id') in ids_inclus]
            ids_inclus.extend([d['id'] for d in directions])
            departements = [d for d in db.get_all_departements() if d.get('direction_id') in ids_inclus]
            ids_inclus.extend([d['id'] for d in departements])
            services = [s for s in db.get_all_services() if s.get('departement_id') in ids_inclus]
            ids_inclus.extend([s['id'] for s in services])
            equipes = [e for e in db.get_all_equipes() if e.get('service_id') in ids_inclus]
            ids_inclus.extend([e['id'] for e in equipes])
        elif entite_type == 'division':
            directions = [d for d in db.get_all_directions() if d.get('division_id') == entite_id]
            ids_inclus.extend([d['id'] for d in directions])
            departements = [d for d in db.get_all_departements() if d.get('direction_id') in ids_inclus]
            ids_inclus.extend([d['id'] for d in departements])
            services = [s for s in db.get_all_services() if s.get('departement_id') in ids_inclus]
            ids_inclus.extend([s['id'] for s in services])
            equipes = [e for e in db.get_all_equipes() if e.get('service_id') in ids_inclus]
            ids_inclus.extend([e['id'] for e in equipes])
        elif entite_type == 'direction':
            departements = [d for d in db.get_all_departements() if d.get('direction_id') == entite_id]
            ids_inclus.extend([d['id'] for d in departements])
            services = [s for s in db.get_all_services() if s.get('departement_id') in ids_inclus]
            ids_inclus.extend([s['id'] for s in services])
            equipes = [e for e in db.get_all_equipes() if e.get('service_id') in ids_inclus]
            ids_inclus.extend([e['id'] for e in equipes])
        elif entite_type == 'departement':
            services = [s for s in db.get_all_services() if s.get('departement_id') == entite_id]
            ids_inclus.extend([s['id'] for s in services])
            equipes = [e for e in db.get_all_equipes() if e.get('service_id') in ids_inclus]
            ids_inclus.extend([e['id'] for e in equipes])
        elif entite_type == 'service':
            equipes = [e for e in db.get_all_equipes() if e.get('service_id') == entite_id]
            ids_inclus.extend([e['id'] for e in equipes])

    # Construire la requête SQL avec les IDs
    # On utilise une requête paramétrée avec un nombre variable de ? (x)
    placeholders = ','.join(['?'] * len(ids_inclus))
    field_map = {
        'activite': 'p.activite_id',
        'division': 'p.division_id',
        'direction': 'p.direction_id',
        'departement': 'p.departement_id',
        'service': 'p.service_id',
        'equipe': 'p.equipe_id'
    }
    colonne = field_map.get(entite_type, 'p.id')
    query = f"""
        SELECT p.id, p.matricule, p.nom, p.prenom, p.fonction, p.statut,
               a.nom_activite as activite_nom,
               d.nom_division as division_nom,
               dir.nom_direction as direction_nom,
               dep.nom_departement as departement_nom,
               s.nom_service as service_nom,
               e.nom_equipe as equipe_nom
        FROM personnel p
        LEFT JOIN activites a ON p.activite_id = a.id
        LEFT JOIN divisions d ON p.division_id = d.id
        LEFT JOIN directions dir ON p.direction_id = dir.id
        LEFT JOIN departements dep ON p.departement_id = dep.id
        LEFT JOIN services s ON p.service_id = s.id
        LEFT JOIN equipes e ON p.equipe_id = e.id
        WHERE {colonne} IN ({placeholders})
        ORDER BY p.nom, p.prenom
    """
    with db.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(query, ids_inclus)
        personnel = [dict(row) for row in cursor.fetchall()]

    # Pagination
    total = len(personnel)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    personnel_page = personnel[offset:offset+per_page]

    filter_args = {'type': entite_type, 'id': entite_id, 'descendants': 'true' if inclure_descendants else 'false'}
    stats = get_system_stats()  # si vous avez cette fonction

    return render_template('personnel/hierarchie_liste.html',
                           personnel=personnel_page,
                           entite_type=entite_type,
                           entite_id=entite_id,
                           nom_entite=nom_entite,
                           inclure_descendants=inclure_descendants,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args,
                           stats=stats)

# -------------------------------------------------------------------
# Gestion hiérarchique (CRUD complet)
# -------------------------------------------------------------------
@app.route('/hierarchie/activites')
@login_required
def liste_activites():
    activites = db.get_all_activites()
    return render_template('hierarchie/activites.html', activites=activites)

@app.route('/hierarchie/activites/ajouter', methods=['POST'])
@login_required
@admin_required
def ajouter_activite():
    nom = request.form.get['nom']
    code = request.form.get('code')
    responsable = request.form.get('responsable')
    if not nom:
        return jsonify({'success': False, 'message': 'Le nom est obligatoire'})
    
    try:
        activite_id = db.add_activite(nom, code, responsable)
        return jsonify({'success': True, 'id': activite_id})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/hierarchie/activites/<int:aid>/modifier', methods=['GET', 'POST'])
@login_required
@admin_required
def modifier_activite(aid):
    activite = db.get_activite(aid)
    if not activite:
        flash("Activité non trouvée.", "danger")
        return redirect(url_for('liste_activites'))
    if request.method == 'POST':
        nom = request.form['nom']
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        active = 'active' in request.form
        
        try:
            db.update_activite(aid, nom, code, responsable, active)
            flash("Activité modifiée avec succès", "success")
            return redirect(url_for('liste_activites'))
        except Exception as e:
            flash(f"Erreur : {str(e)}", "danger")
    return render_template('hierarchie/modifier_activite.html', activite=activite)

@app.route('/hierarchie/activites/<int:aid>/supprimer', methods=['POST'])
@login_required
@admin_required
def supprimer_activite(aid):
    if db.delete_activite(aid):
        flash("Activité supprimée.", "success")
    else:
        flash("Impossible de supprimer : cette activité est utilisée.", "danger")
    return redirect(url_for('liste_activites'))

@app.route('/hierarchie/divisions')
@login_required
def liste_divisions():
    divisions = db.get_all_divisions()
    activites = db.get_all_activites()
    return render_template('hierarchie/divisions.html', divisions=divisions, activites=activites)

#-----------------------------------------------------------------------------
# Ajout d'une division avec gestion des erreurs et validation
#-----------------------------------------------------------------------------
@app.route('/hierarchie/division/ajouter', methods=['POST'])
@login_required
def ajouter_division():
    try:
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        activite_id = request.form.get('activite_id') or None
        description = request.form.get('description')
        
        if not nom:
            return jsonify({'success': False, 'message': 'Le nom est obligatoire'})
        
        division_id = db.add_division(nom, code, responsable, description, activite_id)
        return jsonify({'success': True, 'id': division_id})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})
    
#-------------------------------------------------------------------------------
# Modification d'une division avec validation et gestion des erreurs
#-------------------------------------------------------------------------------

@app.route('/hierarchie/divisions/<int:did>/modifier', methods=['GET', 'POST'])
@login_required
def modifier_division(did):
    division = db.get_division(did)
    if not division:
        flash("Division non trouvée", "danger")
        return redirect(url_for('liste_divisions'))
    
    activites = db.get_all_activites()
    
    if request.method == 'POST':
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        activite_id = request.form.get('activite_id') or None
        description = request.form.get('description')
        active = 'active' in request.form
        
        try:
            db.update_division(did, nom, code, responsable, activite_id, description, active)
            flash("Division modifiée avec succès", "success")
            return redirect(url_for('liste_divisions'))
        except Exception as e:
            flash(f"Erreur : {str(e)}", "danger")
    
    return render_template('hierarchie/modifier_division.html', 
                           division=division, 
                           activites=activites)
#-------------------------------------------------------------------------------
# Suppression d'une division avec gestion des erreurs et retour JSON
#-------------------------------------------------------------------------------
@app.route('/hierarchie/division/<int:did>/supprimer', methods=['POST'])
@login_required
def supprimer_division(did):
    try:
        db.delete_division(did)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ============================================================
# GESTION DES DIRECTIONS
# ============================================================

@app.route('/hierarchie/directions')
@login_required
def liste_directions():
    directions = db.get_all_directions()
    divisions = db.get_all_divisions()
    return render_template('hierarchie/directions.html', directions=directions, divisions=divisions)

@app.route('/hierarchie/direction/ajouter', methods=['POST'])
@login_required
def ajouter_direction():
    try:
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        division_id = request.form.get('division_id') or None
        description = request.form.get('description')
        
        if not nom:
            return jsonify({'success': False, 'message': 'Le nom est obligatoire'})
        
        direction_id = db.add_direction(nom, code, responsable, description, division_id)
        return jsonify({'success': True, 'id': direction_id})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/hierarchie/directions/<int:did>/modifier', methods=['GET', 'POST'])
@login_required
@admin_required
def modifier_direction(did):
    direction = db.get_direction(did)
    if not direction:
        flash("Direction non trouvée.", "danger")
        return redirect(url_for('liste_directions'))
    divisions = db.get_all_divisions()
    if request.method == 'POST':
        nom = request.form['nom']
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        description = request.form.get('description')
        division_id = request.form.get('division_id') or None
        active = 'active' in request.form
        try:
            db.update_direction(did, nom, code, responsable, division_id, description, active)
            flash("Direction modifiée avec succès", "success")
            return redirect(url_for('liste_directions'))
        except Exception as e:
            flash(f"Erreur : {str(e)}", "danger")
    
    return render_template('hierarchie/modifier_direction.html', direction=direction, divisions=divisions)

@app.route('/hierarchie/direction/<int:did>/supprimer', methods=['POST'])
@login_required
def supprimer_direction(did):
    try:
        db.delete_direction(did)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})
    
    
# ============================================================
# GESTION DES DÉPARTEMENTS
# ============================================================

@app.route('/hierarchie/departements')
@login_required
def liste_departements():
    departements = db.get_all_departements()
    directions = db.get_all_directions()
    return render_template('hierarchie/departements.html', departements=departements, directions=directions)

@app.route('/hierarchie/departement/ajouter', methods=['POST'])
@login_required
def ajouter_departement():
    try:
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        direction_id = request.form.get('direction_id') or None
        description = request.form.get('description')
        
        if not nom:
            return jsonify({'success': False, 'message': 'Le nom est obligatoire'})
        
        departement_id = db.add_departement(nom, code, responsable, description, direction_id)
        return jsonify({'success': True, 'id': departement_id})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/hierarchie/departements/<int:did>/modifier', methods=['GET', 'POST'])
@login_required
def modifier_departement(did):
    departement = db.get_departement(did)
    if not departement:
        flash("Département non trouvé", "danger")
        return redirect(url_for('liste_departements'))
    
    directions = db.get_all_directions()
    
    if request.method == 'POST':
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        direction_id = request.form.get('direction_id') or None
        description = request.form.get('description')
        active = 'active' in request.form
        
        try:
            db.update_departement(did, nom, code, responsable, direction_id, description, active)
            flash("Département modifié avec succès", "success")
            return redirect(url_for('liste_departements'))
        except Exception as e:
            flash(f"Erreur : {str(e)}", "danger")
    
    return render_template('hierarchie/modifier_departement.html', 
                           departement=departement, 
                           directions=directions)

@app.route('/hierarchie/departement/<int:did>/supprimer', methods=['POST'])
@login_required
def supprimer_departement(did):
    try:
        db.delete_departement(did)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ============================================================
# GESTION DES SERVICES
# ============================================================

@app.route('/hierarchie/services')
@login_required
def liste_services():
    services = db.get_all_services()
    departements = db.get_all_departements()
    return render_template('hierarchie/services.html', services=services, departements=departements)

@app.route('/hierarchie/service/ajouter', methods=['POST'])
@login_required
def ajouter_service():
    try:
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        departement_id = request.form.get('departement_id') or None
        description = request.form.get('description')
        
        if not nom:
            return jsonify({'success': False, 'message': 'Le nom est obligatoire'})
        
        service_id = db.add_service(nom, code, responsable, description, departement_id)
        return jsonify({'success': True, 'id': service_id})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/hierarchie/services/<int:sid>/modifier', methods=['GET', 'POST'])
@login_required
def modifier_service(sid):
    service = db.get_service(sid)
    if not service:
        flash("Service non trouvé", "danger")
        return redirect(url_for('liste_services'))
    
    departements = db.get_all_departements()
    
    if request.method == 'POST':
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        departement_id = request.form.get('departement_id') or None
        description = request.form.get('description')
        active = 'active' in request.form
        
        try:
            db.update_service(sid, nom, code, responsable, departement_id, description, active)
            flash("Service modifié avec succès", "success")
            return redirect(url_for('liste_services'))
        except Exception as e:
            flash(f"Erreur : {str(e)}", "danger")
    
    return render_template('hierarchie/modifier_service.html', 
                           service=service, 
                           departements=departements)

@app.route('/hierarchie/service/<int:sid>/supprimer', methods=['POST'])
@login_required
def supprimer_service(sid):
    try:
        db.delete_service(sid)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# ============================================================
# GESTION DES ÉQUIPES
# ============================================================

@app.route('/hierarchie/equipes')
@login_required
def liste_equipes():
    equipes = db.get_all_equipes()
    services = db.get_all_services()
    return render_template('hierarchie/equipes.html', equipes=equipes, services=services)

@app.route('/hierarchie/equipe/ajouter', methods=['POST'])
@login_required
def ajouter_equipe():
    try:
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        service_id = request.form.get('service_id') or None
        description = request.form.get('description')
        
        if not nom:
            return jsonify({'success': False, 'message': 'Le nom est obligatoire'})
        
        equipe_id = db.add_equipe(nom, code, responsable, description, service_id)
        return jsonify({'success': True, 'id': equipe_id})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/hierarchie/equipes/<int:eid>/modifier', methods=['GET', 'POST'])
@login_required
def modifier_equipe(eid):
    equipe = db.get_equipe(eid)
    if not equipe:
        flash("Équipe non trouvée", "danger")
        return redirect(url_for('liste_equipes'))
    
    services = db.get_all_services()
    
    if request.method == 'POST':
        nom = request.form.get('nom')
        code = request.form.get('code')
        responsable = request.form.get('responsable')
        service_id = request.form.get('service_id') or None
        description = request.form.get('description')
        active = 'active' in request.form
        
        try:
            db.update_equipe(eid, nom, code, responsable, service_id, description, active)
            flash("Équipe modifiée avec succès", "success")
            return redirect(url_for('liste_equipes'))
        except Exception as e:
            flash(f"Erreur : {str(e)}", "danger")
    
    return render_template('hierarchie/modifier_equipe.html', 
                           equipe=equipe, 
                           services=services)

@app.route('/hierarchie/equipe/<int:eid>/supprimer', methods=['POST'])
@login_required
def supprimer_equipe(eid):
    try:
        db.delete_equipe(eid)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


# -------------------------------------------------------------------
# Routes API pour les menus déroulants dynamiques
# -------------------------------------------------------------------
@app.route('/api/agent/<int:agent_id>')
@login_required
def api_get_agent(agent_id):
    agent = db.get_personnel(agent_id)
    if not agent:
        return jsonify({'error': 'Agent non trouvé'}), 404
    
    return jsonify({
        'id': agent['id'],
        'matricule': agent['matricule'],
        'nom': agent['nom'],
        'prenom': agent['prenom'],
        'fonction': agent.get('fonction', ''),
        'division_nom': agent.get('division_nom', ''),
        'departement_nom': agent.get('departement_nom', ''),
        'service_nom': agent.get('service_nom', ''),
        'equipe_nom': agent.get('equipe_nom', ''),
        'direction_nom': agent.get('direction_nom', ''),
        'quart_nom': agent.get('quart_nom', '')
    })


@app.route('/api/divisions/<int:activite_id>')
@login_required
def api_divisions(activite_id):
    with db.get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT id, nom_division FROM divisions WHERE activite_id = ? AND active = 1 ORDER BY nom_division", (activite_id,))
        return jsonify([{'id': r[0], 'nom': r[1]} for r in c.fetchall()])

@app.route('/api/directions/<int:division_id>')
@login_required
def api_directions(division_id):
    with db.get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT id, nom_direction FROM directions WHERE division_id = ? AND active = 1 ORDER BY nom_direction", (division_id,))
        return jsonify([{'id': r[0], 'nom': r[1]} for r in c.fetchall()])

@app.route('/api/departements/<int:direction_id>')
@login_required
def api_departements(direction_id):
    with db.get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT id, nom_departement FROM departements WHERE direction_id = ? AND active = 1 ORDER BY nom_departement", (direction_id,))
        return jsonify([{'id': r[0], 'nom': r[1]} for r in c.fetchall()])

@app.route('/api/services/<int:departement_id>')
@login_required
def api_services(departement_id):
    with db.get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT id, nom_service FROM services WHERE departement_id = ? AND active = 1 ORDER BY nom_service", (departement_id,))
        return jsonify([{'id': r[0], 'nom': r[1]} for r in c.fetchall()])

@app.route('/api/equipes/<int:service_id>')
@login_required
def api_equipes(service_id):
    with db.get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT id, nom_equipe FROM equipes WHERE service_id = ? AND active = 1 ORDER BY nom_equipe", (service_id,))
        return jsonify([{'id': r[0], 'nom': r[1]} for r in c.fetchall()])

# -------------------------------------------------------------------
# Rapports (présence, absences, retards) – à compléter selon vos besoins
# -------------------------------------------------------------------
# (Les routes de rapports sont supposées exister ailleurs. Si ce n'est pas le cas,
#  vous devrez les ajouter en vous inspirant des classes Tkinter.)
# -------------------------------------------------------------------
# Gestion des quarts
# -------------------------------------------------------------------
@app.route('/quarts')
@login_required
def liste_quarts():
    quarts = db.get_all_quarts()
    return render_template('quarts/liste.html', quarts=quarts)

@app.route('/quarts/ajouter', methods=['POST'])
@login_required
@admin_required
def ajouter_quart():
    nom = request.form['nom']
    debut = request.form['heure_debut']
    fin = request.form['heure_fin']
    nuit = 1 if 'est_nuit' in request.form else 0
    description = request.form.get('description')
    db.add_quart(nom, debut, fin, nuit, None, description)
    flash(f"Quart '{nom}' ajouté.", "success")
    return redirect(url_for('liste_quarts'))

@app.route('/quarts/<int:qid>/modifier', methods=['GET', 'POST'])
@login_required
@admin_required
def modifier_quart(qid):
    quart = db.get_quart(qid)
    if not quart:
        flash("Quart non trouvé.", "danger")
        return redirect(url_for('liste_quarts'))
    if request.method == 'POST':
        nom = request.form['nom']
        heure_debut = request.form['heure_debut']
        heure_fin = request.form['heure_fin']
        est_nuit = 'est_nuit' in request.form
        duree_heures = request.form.get('duree_heures')
        if duree_heures:
            duree_heures = float(duree_heures)
        description = request.form.get('description')
        couleur = request.form.get('couleur')
        active = 'active' in request.form
        db.update_quart(qid, nom, heure_debut, heure_fin, est_nuit, duree_heures, description, couleur, active)
        flash("Quart modifié.", "success")
        return redirect(url_for('liste_quarts'))
    return render_template('quarts/modifier.html', quart=quart)

@app.route('/quarts/<int:qid>/supprimer', methods=['POST'])
@login_required
@admin_required
def supprimer_quart(qid):
    if db.delete_quart(qid):
        flash("Quart supprimé.", "success")
    else:
        flash("Impossible de supprimer : ce quart est utilisé par du personnel.", "danger")
    return redirect(url_for('liste_quarts'))

# -------------------------------------------------------------------
# Gestion des congés
# -------------------------------------------------------------------
@app.route('/conges')
@login_required
def liste_conges():
    filtres = {}
    statut = request.args.get('statut', 'Tous')
    if statut != 'Tous':
        filtres['statut'] = statut
    matricule = request.args.get('matricule', '')
    if matricule:
        filtres['matricule'] = matricule
    type_conge = request.args.get('type_conge', 'Tous')
    if type_conge != 'Tous':
        filtres['type_conge'] = type_conge
    conges = db.get_conges(filtres)
    return render_template('conges/liste.html', conges=conges, statut=statut, matricule=matricule, type_conge=type_conge)

@app.route('/conges/ajouter', methods=['GET', 'POST'])
@login_required
def ajouter_conge():
    if request.method == 'POST':
        agent_id = request.form.get('personnel_id')
        type_conge = request.form.get('type_conge')
        date_debut = request.form.get('date_debut')
        date_fin = request.form.get('date_fin')
        motif = request.form.get('motif')
        justificatif = request.form.get('justificatif')
        agent = db.get_personnel(agent_id)
        if not agent:
            flash("Agent non trouvé", "danger")
            return redirect(url_for('ajouter_conge'))
        debut = datetime.strptime(date_debut, '%Y-%m-%d').date()
        fin = datetime.strptime(date_fin, '%Y-%m-%d').date()
        duree = (fin - debut).days + 1
        data = (
            agent_id, agent['matricule'], type_conge, date_debut, date_fin, duree,
            motif, justificatif, 'en_attente', date.today().strftime('%Y-%m-%d')
        )
        db.add_conge(data)
        flash("Demande de congé soumise", "success")
        return redirect(url_for('liste_conges'))
    agents = db.get_personnel()
    return render_template('conges/form.html', agents=agents, conge=None)

@app.route('/conges/<int:cid>/modifier', methods=['GET', 'POST'])
@login_required
def modifier_conge(cid):
    conge = db.get_conge(cid)
    if not conge:
        flash("Congé non trouvé", "danger")
        return redirect(url_for('liste_conges'))
    if request.method == 'POST':
        type_conge = request.form.get('type_conge')
        date_debut = request.form.get('date_debut')
        date_fin = request.form.get('date_fin')
        motif = request.form.get('motif')
        justificatif = request.form.get('justificatif')
        debut = datetime.strptime(date_debut, '%Y-%m-%d').date()
        fin = datetime.strptime(date_fin, '%Y-%m-%d').date()
        duree = (fin - debut).days + 1
        data = (type_conge, date_debut, date_fin, duree, motif, justificatif)
        db.update_conge(cid, data)
        flash("Congé modifié", "success")
        return redirect(url_for('liste_conges'))
    agents = db.get_personnel()
    return render_template('conges/form.html', conge=conge, agents=agents)

@app.route('/conges/<int:cid>/supprimer', methods=['POST'])
@login_required
def supprimer_conge(cid):
    if session['role'] not in ['admin', 'superviseur']:
        flash("Non autorisé", "danger")
        return redirect(url_for('liste_conges'))
    db.delete_conge(cid)
    flash("Congé supprimé", "success")
    return redirect(url_for('liste_conges'))

@app.route('/conges/<int:cid>/approuver', methods=['POST'])
@login_required
def approuver_conge(cid):
    if session['role'] not in ['admin', 'superviseur']:
        flash("Non autorisé", "danger")
        return redirect(url_for('liste_conges'))
    db.update_conge_statut(cid, 'approuve', session['username'])
    flash("Congé approuvé", "success")
    return redirect(url_for('liste_conges'))

@app.route('/conges/<int:cid>/refuser', methods=['POST'])
@login_required
def refuser_conge(cid):
    if session['role'] not in ['admin', 'superviseur']:
        flash("Non autorisé", "danger")
        return redirect(url_for('liste_conges'))
    db.update_conge_statut(cid, 'refuse', session['username'])
    flash("Congé refusé", "success")
    return redirect(url_for('liste_conges'))

# -------------------------------------------------------------------
# Gestion des heures supplémentaires
# -------------------------------------------------------------------
@app.route('/heures_sup')
@login_required
def liste_heures_sup():
    filtres = {}
    statut = request.args.get('statut', 'Tous')
    if statut != 'Tous':
        filtres['statut'] = statut
    matricule = request.args.get('matricule', '')
    if matricule:
        filtres['matricule'] = matricule
    date_debut = request.args.get('date_debut', '')
    if date_debut:
        filtres['date_debut'] = date_debut
    date_fin = request.args.get('date_fin', '')
    if date_fin:
        filtres['date_fin'] = date_fin
    type_heure_sup = request.args.get('type_heure_sup', 'Tous')
    if type_heure_sup != 'Tous':
        filtres['type_heure_sup'] = type_heure_sup
    heures = db.get_heures_supplementaires(filtres)
    return render_template('heures_sup/liste.html', heures=heures, statut=statut, matricule=matricule,
                           date_debut=date_debut, date_fin=date_fin, type_heure_sup=type_heure_sup)

@app.route('/heures_sup/ajouter', methods=['GET', 'POST'])
@login_required
def ajouter_heure_sup():
    if request.method == 'POST':
        agent_id = request.form.get('personnel_id')
        date_debut = request.form.get('date_heure_debut')
        date_fin = request.form.get('date_heure_fin')
        type_heure_sup = request.form.get('type_heure_sup')
        motif = request.form.get('motif')
        agent = db.get_personnel(agent_id)
        if not agent:
            flash("Agent non trouvé", "danger")
            return redirect(url_for('ajouter_heure_sup'))
        dt_debut = datetime.strptime(date_debut, '%Y-%m-%dT%H:%M')
        dt_fin = datetime.strptime(date_fin, '%Y-%m-%dT%H:%M')
        duree = (dt_fin - dt_debut).total_seconds() / 3600
        # Détection weekend/ferié (à améliorer)
        est_weekend = 1 if dt_debut.weekday() >= 5 else 0
        est_ferie = 0  # à vérifier
        taux = {'normale': 1.25, 'nuit': 1.5, 'weekend': 1.75, 'jour_ferie': 2.0}.get(type_heure_sup, 1.25)
        data = (
            agent_id, agent['matricule'], date_debut, date_fin, duree,
            type_heure_sup, taux, est_weekend, est_ferie, motif, 'en_attente'
        )
        db.add_heure_sup(data)
        flash("Demande d'heures supplémentaires soumise", "success")
        return redirect(url_for('liste_heures_sup'))
    agents = db.get_personnel()
    return render_template('heures_sup/form.html', agents=agents, heure=None)

@app.route('/heures_sup/<int:hid>/modifier', methods=['GET', 'POST'])
@login_required
def modifier_heure_sup(hid):
    heure = db.get_heure_sup(hid)
    if not heure:
        flash("Heure sup non trouvée", "danger")
        return redirect(url_for('liste_heures_sup'))
    if request.method == 'POST':
        date_debut = request.form.get('date_heure_debut')
        date_fin = request.form.get('date_heure_fin')
        type_heure_sup = request.form.get('type_heure_sup')
        motif = request.form.get('motif')
        dt_debut = datetime.strptime(date_debut, '%Y-%m-%dT%H:%M')
        dt_fin = datetime.strptime(date_fin, '%Y-%m-%dT%H:%M')
        duree = (dt_fin - dt_debut).total_seconds() / 3600
        data = (date_debut, date_fin, duree, type_heure_sup, motif)
        db.update_heure_sup(hid, data)
        flash("Heure sup modifiée", "success")
        return redirect(url_for('liste_heures_sup'))
    agents = db.get_personnel()
    return render_template('heures_sup/form.html', heure=heure, agents=agents)

@app.route('/heures_sup/<int:hid>/supprimer', methods=['POST'])
@login_required
def supprimer_heure_sup(hid):
    if session['role'] not in ['admin', 'superviseur']:
        flash("Non autorisé", "danger")
        return redirect(url_for('liste_heures_sup'))
    db.delete_heure_sup(hid)
    flash("Heure sup supprimée", "success")
    return redirect(url_for('liste_heures_sup'))

@app.route('/heures_sup/<int:hid>/approuver', methods=['POST'])
@login_required
def approuver_heure_sup(hid):
    if session['role'] not in ['admin', 'superviseur']:
        flash("Non autorisé", "danger")
        return redirect(url_for('liste_heures_sup'))
    db.update_heure_sup_statut(hid, 'approuve', session['username'])
    flash("Heure sup approuvée", "success")
    return redirect(url_for('liste_heures_sup'))

@app.route('/heures_sup/<int:hid>/refuser', methods=['POST'])
@login_required
def refuser_heure_sup(hid):
    if session['role'] not in ['admin', 'superviseur']:
        flash("Non autorisé", "danger")
        return redirect(url_for('liste_heures_sup'))
    db.update_heure_sup_statut(hid, 'refuse', session['username'])
    flash("Heure sup refusée", "success")
    return redirect(url_for('liste_heures_sup'))

# ============================================================
# GESTION DES JOURS FÉRIÉS
# ============================================================

@app.route('/liste/jours_feries')
@login_required
def liste_jours_feries():
    jours = db.get_all_jours_feries()
    return render_template('jours_feries/liste.html', jours=jours)

@app.route('/liste/jours_feries/ajouter', methods=['POST'])
@login_required
def ajouter_jour_ferie():
    date_jour = request.form.get('date')
    nom = request.form.get('nom')
    description = request.form.get('description')
    if not date_jour or not nom:
        flash("La date et le nom sont obligatoires", "danger")
        return redirect(url_for('liste_jours_feries'))
    try:
        db.add_jour_ferie(date_jour, nom, description)
        flash("Jour férié ajouté", "success")
    except sqlite3.IntegrityError:
        flash("Cette date existe déjà", "danger")
    except Exception as e:
        flash(f"Erreur : {e}", "danger")
    return redirect(url_for('liste_jours_feries'))

@app.route('/liste/jours_feries/<int:jid>/supprimer', methods=['POST'])
@login_required
def supprimer_jour_ferie(jid):
    try:
        db.delete_jour_ferie(jid)
        flash("Jour férié supprimé", "success")
    except Exception as e:
        flash(f"Erreur : {e}", "danger")
    return redirect(url_for('liste_jours_feries'))

# -------------------------------------------------------------------
# Gestion des retards cumulés
# -------------------------------------------------------------------
@app.route('/retards')
@login_required
def liste_retards():
    mois = request.args.get('mois', datetime.now().strftime('%Y-%m'))
    retards = db.get_retards_cumules(mois=mois)
    return render_template('retards/liste.html', retards=retards, mois=mois)

@app.route('/retards/<int:rid>/justifier', methods=['POST'])
@login_required
def justifier_retard(rid):
    justification = request.form.get('justification', '')
    if db.justifier_retard(rid, justification):
        flash("Retard justifié.", "success")
    else:
        flash("Erreur lors de la justification.", "danger")
    return redirect(url_for('liste_retards'))

# -------------------------------------------------------------------
# Styles PDF personnalisés (défini avant son utilisation)
# -------------------------------------------------------------------
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm

def get_pdf_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='CustomTitle',
        parent=styles['Title'],
        fontSize=16,
        textColor=colors.HexColor('#2c3e50'),
        alignment=1,
        spaceAfter=20
    ))
    styles.add(ParagraphStyle(
        name='CustomHeader',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.white,
        alignment=1,
        backColor=colors.HexColor('#3498db')
    ))
    return styles

# -------------------------------------------------------------------
# Générateur de jours ouvrés (exclut weekends et jours fériés) - utilisé pour les rapports et les calculs de congés
# -------------------------------------------------------------------
def jours_ouvres(debut, fin):
    jour = debut
    while jour <= fin:
        if not db.est_weekend(jour) and not db.est_jour_ferie(jour):
            yield jour
        jour += timedelta(days=1)

#-------------------------------------------------------------------
# Rapport de présence
#-------------------------------------------------------------------
@app.route('/rapports/presence', methods=['GET', 'POST'])
@login_required
def rapport_presence():
    if request.method == 'POST':
        date_debut = request.form.get('date_debut')
        date_fin = request.form.get('date_fin')
        search = request.form.get('search', '').strip()
        direction = request.form.get('direction')
        departement = request.form.get('departement')
        service = request.form.get('service')
        equipe = request.form.get('equipe')
        type_rapport = request.form.get('type_rapport', 'resume')
        
        # Stocker les filtres en session
        session['presence_filters'] = {
            'date_debut': date_debut,
            'date_fin': date_fin,
            'search': search,
            'direction': direction,
            'departement': departement,
            'service': service,
            'equipe': equipe,
            'type_rapport': type_rapport
        }
    else:
        filters = session.get('presence_filters', {})
        date_debut = filters.get('date_debut', date.today().replace(day=1).strftime('%Y-%m-%d'))
        date_fin = filters.get('date_fin', date.today().strftime('%Y-%m-%d'))
        search = filters.get('search', '')
        direction = filters.get('direction', '')
        departement = filters.get('departement', '')
        service = filters.get('service', '')
        equipe = filters.get('equipe', '')
        type_rapport = filters.get('type_rapport', 'resume')
    
    filter_args = {
        'date_debut': date_debut, 'date_fin': date_fin, 'search': search,
        'direction': direction, 'departement': departement,
        'service': service, 'equipe': equipe, 'type_rapport': type_rapport
    }
    filter_args = {k: v for k, v in filter_args.items() if v}
    
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    if type_rapport == 'resume':
        resultats = db.get_presence_report(
            date_debut=date_debut, date_fin=date_fin, search=search,
            direction_nom=direction, departement_nom=departement,
            service_nom=service, equipe_nom=equipe, type_rapport='resume'
        )
        total = len(resultats)
        total_pages = (total + per_page - 1) // per_page if total > 0 else 1
        offset = (page - 1) * per_page
        resultats_page = resultats[offset:offset+per_page]
    else:
        details = db.get_presence_report(
            date_debut=date_debut, date_fin=date_fin, search=search,
            direction_nom=direction, departement_nom=departement,
            service_nom=service, equipe_nom=equipe, type_rapport='detail'
        )
        total = len(details)
        total_pages = (total + per_page - 1) // per_page if total > 0 else 1
        offset = (page - 1) * per_page
        resultats_page = details[offset:offset+per_page]
    
    # Charger les listes pour les filtres
    directions = db.get_all_directions()
    departements = db.get_all_departements()
    services = db.get_all_services()
    equipes = db.get_all_equipes()
    
    if type_rapport == 'resume':
        return render_template('rapports/presence_resultats_resume.html',
                               resultats=resultats_page, page=page,
                               total_pages=total_pages, per_page=per_page,
                               filter_args=filter_args, date_debut=date_debut,
                               date_fin=date_fin, search=search,
                               direction=direction, departement=departement,
                               service=service, equipe=equipe,
                               type_rapport=type_rapport,
                               directions=directions, departements=departements,
                               services=services, equipes=equipes)
    else:
        return render_template('rapports/presence_resultats_detail.html',
                               details=resultats_page, page=page,
                               total_pages=total_pages, per_page=per_page,
                               filter_args=filter_args, date_debut=date_debut,
                               date_fin=date_fin, search=search,
                               direction=direction, departement=departement,
                               service=service, equipe=equipe,
                               type_rapport=type_rapport,
                               directions=directions, departements=departements,
                               services=services, equipes=equipes)
#-------------------------------------------------------------------
# Rapport d'absences
#-------------------------------------------------------------------
@app.route('/rapports/absences', methods=['GET', 'POST'])
@login_required
def rapport_absences():
    if request.method == 'POST':
        date_debut = request.form.get('date_debut')
        date_fin = request.form.get('date_fin')
        search = request.form.get('search', '').strip()
        direction = request.form.get('direction')
        departement = request.form.get('departement')
        service = request.form.get('service')
        equipe = request.form.get('equipe')
        type_rapport = request.form.get('type_rapport', 'resume')
        
        session['absences_filters'] = {
            'date_debut': date_debut, 'date_fin': date_fin, 'search': search,
            'direction': direction, 'departement': departement,
            'service': service, 'equipe': equipe, 'type_rapport': type_rapport
        }
    else:
        filters = session.get('absences_filters', {})
        date_debut = filters.get('date_debut', date.today().replace(day=1).strftime('%Y-%m-%d'))
        date_fin = filters.get('date_fin', date.today().strftime('%Y-%m-%d'))
        search = filters.get('search', '')
        direction = filters.get('direction', '')
        departement = filters.get('departement', '')
        service = filters.get('service', '')
        equipe = filters.get('equipe', '')
        type_rapport = filters.get('type_rapport', 'resume')
    
    filter_args = {
        'date_debut': date_debut, 'date_fin': date_fin, 'search': search,
        'direction': direction, 'departement': departement,
        'service': service, 'equipe': equipe, 'type_rapport': type_rapport
    }
    filter_args = {k: v for k, v in filter_args.items() if v}
    
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    if type_rapport == 'resume':
        resultats = db.get_absence_report(
            date_debut=date_debut, date_fin=date_fin, search=search,
            direction_nom=direction, departement_nom=departement,
            service_nom=service, equipe_nom=equipe, type_rapport='resume'
        )
        total = len(resultats)
        total_pages = (total + per_page - 1) // per_page if total > 0 else 1
        offset = (page - 1) * per_page
        resultats_page = resultats[offset:offset+per_page]
    else:
        details = db.get_absence_report(
            date_debut=date_debut, date_fin=date_fin, search=search,
            direction_nom=direction, departement_nom=departement,
            service_nom=service, equipe_nom=equipe, type_rapport='detail'
        )
        total = len(details)
        total_pages = (total + per_page - 1) // per_page if total > 0 else 1
        offset = (page - 1) * per_page
        resultats_page = details[offset:offset+per_page]
    
    directions = db.get_all_directions()
    departements = db.get_all_departements()
    services = db.get_all_services()
    equipes = db.get_all_equipes()
    
    if type_rapport == 'resume':
        return render_template('rapports/absences_resultats_resume.html',
                               resultats=resultats_page, page=page,
                               total_pages=total_pages, per_page=per_page,
                               filter_args=filter_args, date_debut=date_debut,
                               date_fin=date_fin, search=search,
                               direction=direction, departement=departement,
                               service=service, equipe=equipe, type_rapport=type_rapport,
                               directions=directions, departements=departements,
                               services=services, equipes=equipes)
    else:
        return render_template('rapports/absences_resultats_detail.html',
                               details=resultats_page, page=page,
                               total_pages=total_pages, per_page=per_page,
                               filter_args=filter_args, date_debut=date_debut,
                               date_fin=date_fin, search=search,
                               direction=direction, departement=departement,
                               service=service, equipe=equipe, type_rapport=type_rapport,
                               directions=directions, departements=departements,
                               services=services, equipes=equipes)
#----------------------------------------------------------------------
#
#---------------------------------------------------------------------
@app.route('/rapports/retards', methods=['GET', 'POST'])
@login_required
def rapport_retards():
    if request.method == 'POST':
        date_debut = request.form.get('date_debut')
        date_fin = request.form.get('date_fin')
        search = request.form.get('search', '').strip()
        direction = request.form.get('direction')
        departement = request.form.get('departement')
        service = request.form.get('service')
        equipe = request.form.get('equipe')
        type_rapport = request.form.get('type_rapport', 'resume')
        
        session['retards_filters'] = {
            'date_debut': date_debut,
            'date_fin': date_fin,
            'search': search,
            'direction': direction,
            'departement': departement,
            'service': service,
            'equipe': equipe,
            'type_rapport': type_rapport
        }
    else:
        filters = session.get('retards_filters', {})
        date_debut = filters.get('date_debut', date.today().replace(day=1).strftime('%Y-%m-%d'))
        date_fin = filters.get('date_fin', date.today().strftime('%Y-%m-%d'))
        search = filters.get('search', '')
        direction = filters.get('direction', '')
        departement = filters.get('departement', '')
        service = filters.get('service', '')
        equipe = filters.get('equipe', '')
        type_rapport = filters.get('type_rapport', 'resume')
    
    # Valeurs par défaut si vides
    if not date_debut:
        date_debut = date.today().replace(day=1).strftime('%Y-%m-%d')
    if not date_fin:
        date_fin = date.today().strftime('%Y-%m-%d')
    
    filter_args = {}
    filter_args['date_debut'] = date_debut
    filter_args['date_fin'] = date_fin
    if search:
        filter_args['search'] = search
    if direction and direction not in ('', 'Toutes'):
        filter_args['direction'] = direction
    if departement and departement not in ('', 'Tous'):
        filter_args['departement'] = departement
    if service and service not in ('', 'Tous'):
        filter_args['service'] = service
    if equipe and equipe not in ('', 'Toutes'):
        filter_args['equipe'] = equipe
    filter_args['type_rapport'] = type_rapport
    
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    if type_rapport == 'resume':
        resultats = db.get_retard_report(
            date_debut=date_debut,
            date_fin=date_fin,
            search=search,
            direction_nom=direction if direction and direction != 'Toutes' else None,
            departement_nom=departement if departement and departement != 'Tous' else None,
            service_nom=service if service and service != 'Tous' else None,
            equipe_nom=equipe if equipe and equipe != 'Toutes' else None,
            type_rapport='resume'
        )
        total = len(resultats)
        total_pages = (total + per_page - 1) // per_page if total > 0 else 1
        offset = (page - 1) * per_page
        resultats_page = resultats[offset:offset+per_page]
    else:
        details = db.get_retard_report(
            date_debut=date_debut,
            date_fin=date_fin,
            search=search,
            direction_nom=direction if direction and direction != 'Toutes' else None,
            departement_nom=departement if departement and departement != 'Tous' else None,
            service_nom=service if service and service != 'Tous' else None,
            equipe_nom=equipe if equipe and equipe != 'Toutes' else None,
            type_rapport='detail'
        )
        total = len(details)
        total_pages = (total + per_page - 1) // per_page if total > 0 else 1
        offset = (page - 1) * per_page
        details_page = details[offset:offset+per_page]
    
    # Charger les listes pour les filtres
    directions = db.get_all_directions()
    departements = db.get_all_departements()
    services = db.get_all_services()
    equipes = db.get_all_equipes()
    
    if type_rapport == 'resume':
        return render_template('rapports/retards_resultats_resume.html',
                               resultats=resultats_page,
                               page=page,
                               total_pages=total_pages,
                               per_page=per_page,
                               filter_args=filter_args,
                               date_debut=date_debut,
                               date_fin=date_fin,
                               search=search,
                               direction=direction,
                               departement=departement,
                               service=service,
                               equipe=equipe,
                               type_rapport=type_rapport,
                               directions=directions,
                               departements=departements,
                               services=services,
                               equipes=equipes)
    else:
        return render_template('rapports/retards_resultats_detail.html',
                               details=details_page,
                               page=page,
                               total_pages=total_pages,
                               per_page=per_page,
                               filter_args=filter_args,
                               date_debut=date_debut,
                               date_fin=date_fin,
                               search=search,
                               direction=direction,
                               departement=departement,
                               service=service,
                               equipe=equipe,
                               type_rapport=type_rapport,
                               directions=directions,
                               departements=departements,
                               services=services,
                               equipes=equipes)
#-------------------------------------------------------------------
# Rapport export PDF générique
#-------------------------------------------------------------------
@app.route('/rapports/export_pdf', methods=['POST'])
@login_required
def exporter_rapport_pdf():
    """
    Export PDF générique (page courante) pour tous les rapports.
    Reçoit : titre, sous_titre, colonnes, lignes.
    """
    from weasyprint import HTML
    from weasyprint.text.fonts import FontConfiguration
    from datetime import datetime

    try:
        data = request.get_json()
        titre = data.get('titre', 'Rapport')
        sous_titre = data.get('sous_titre', '')
        colonnes = data.get('colonnes', [])
        lignes = data.get('lignes', [])

        if not lignes:
            return jsonify({'error': 'Aucune donnée'}), 400

        table_html = '<table border="1" style="border-collapse:collapse; width:100%"><thead><tr style="background:#3498db;color:white">'
        for col in colonnes:
            table_html += f'<th style="padding:8px">{col}</th>'
        table_html += '</tr></thead><tbody>'
        for ligne in lignes:
            table_html += '<tr>'
            for cell in ligne:
                table_html += f'<td style="padding:6px">{cell}</td>'
            table_html += '</tr>'
        table_html += '</tbody></table>'

        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head><meta charset="UTF-8"><title>{titre}</title>
        <style>
            @page {{ size: landscape; margin: 1.5cm; }}
            body {{ font-family: DejaVu Sans, Arial, sans-serif; font-size: 10pt; }}
            h1 {{ text-align: center; }}
            table {{ width: 100%; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 6px; text-align: left; }}
            th {{ background-color: #3498db; color: white; }}
        </style>
        </head>
        <body>
            <h1>{titre}</h1>
            <p>{sous_titre}</p>
            {table_html}
            <p style="text-align:center; margin-top:30px; font-size:8pt;">Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </body>
        </html>
        """
        font_config = FontConfiguration()
        pdf = HTML(string=full_html).write_pdf(font_config=font_config)
        response = make_response(pdf)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename=rapport_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        return response
    except Exception as e:
        return jsonify({'error': str(e)}), 500
#---------------------------------------------------------------------------
# Styles PDF personnalisés (défini avant son utilisation)
#---------------------------------------------------------------------------
def get_pdf_styles():
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CustomTitle',
                              parent=styles['Title'],
                              fontSize=16,
                              textColor=colors.HexColor('#2c3e50'),
                              alignment=1,
                              spaceAfter=20))
    return styles

#---------------------------------------------------------------------------
# Statistiques
#-----------------------------------------------------------------------
@app.route('/statistiques/retards_cumules', methods=['GET', 'POST'])
@login_required
def statistiques_retards_cumules():
    """Statistiques cumulées des retards par jour sur une période donnée"""
    
    # Dates par défaut : mois en cours
    today = date.today()
    default_date_debut = today.replace(day=1)
    default_date_fin = today
    
    if request.method == 'POST':
        date_debut = datetime.strptime(request.form.get('date_debut'), '%Y-%m-%d').date()
        date_fin = datetime.strptime(request.form.get('date_fin'), '%Y-%m-%d').date()
        direction = request.form.get('direction')
        departement = request.form.get('departement')
        service = request.form.get('service')
        equipe = request.form.get('equipe')
        search = request.form.get('search', '').strip()
    else:
        date_debut = default_date_debut
        date_fin = default_date_fin
        direction = request.args.get('direction', '')
        departement = request.args.get('departement', '')
        service = request.args.get('service', '')
        equipe = request.args.get('equipe', '')
    
    # Pagination
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    
    # Récupérer les statistiques
    stats = db.get_retards_cumules_par_jour(
        date_debut=date_debut,
        date_fin=date_fin,
        direction_nom=direction if direction not in ('', 'Toutes') else None,
        departement_nom=departement if departement not in ('', 'Tous') else None,
        service_nom=service if service not in ('', 'Tous') else None,
        equipe_nom=equipe if equipe not in ('', 'Toutes') else None
    )
    
    # Pagination
    total = len(stats)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    stats_page = stats[offset:offset+per_page]
    
    # Calculer les totaux
    total_retards = sum(s['total_retards'] for s in stats)
    total_minutes = sum(s['total_minutes'] for s in stats)
    total_agents = len(set(s['matricule'] for s in stats))
    
    # Filtrer par jour (grouper par date)
    stats_par_jour = {}
    for s in stats:
        if s['date'] not in stats_par_jour:
            stats_par_jour[s['date']] = {
                'date': s['date'],
                'nb_retards': 0,
                'nb_agents': 0,
                'total_minutes': 0
            }
        stats_par_jour[s['date']]['nb_retards'] += s['nb_retards']
        stats_par_jour[s['date']]['total_minutes'] += s['total_minutes']
        stats_par_jour[s['date']]['nb_agents'] += 1
    
    stats_par_jour_liste = sorted(stats_par_jour.values(), key=lambda x: x['date'])
    
    # Charger les listes pour les filtres
    directions = db.get_all_directions()
    departements = db.get_all_departements()
    services = db.get_all_services()
    equipes = db.get_all_equipes()
    
    # Filtres pour la pagination
    filter_args = {
        'date_debut': date_debut.strftime('%Y-%m-%d'),
        'date_fin': date_fin.strftime('%Y-%m-%d'),
        'direction': direction,
        'departement': departement,
        'service': service,
        'equipe': equipe
    }
    filter_args = {k: v for k, v in filter_args.items() if v}
    
    return render_template('statistiques/retards_cumules.html',
                           stats=stats_page,
                           stats_par_jour=stats_par_jour_liste,
                           date_debut=date_debut.strftime('%Y-%m-%d'),
                           date_fin=date_fin.strftime('%Y-%m-%d'),
                           direction=direction,
                           departement=departement,
                           service=service,
                           equipe=equipe,
                           total_retards=total_retards,
                           total_minutes=total_minutes,
                           total_agents=total_agents,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args,
                           directions=directions,
                           departements=departements,
                           services=services,
                           equipes=equipes)

# -------------------------------------------------------------------
# Gestion des utilisateurs (admin)
# -------------------------------------------------------------------
@app.route('/utilisateurs')
@login_required
@admin_required
def liste_utilisateurs():
    utilisateurs = db.get_all_users()
    return render_template('utilisateurs/liste.html', utilisateurs=utilisateurs)

@app.route('/utilisateurs/ajouter', methods=['GET', 'POST'])
@login_required
@admin_required
def ajouter_utilisateur():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        nom = request.form['nom']
        prenom = request.form['prenom']
        email = request.form.get('email')
        role = request.form['role']
        try:
            db.create_user(username, password, nom, prenom, email, role)
            flash("Utilisateur créé avec succès.", "success")
            return redirect(url_for('liste_utilisateurs'))
        except Exception as e:
            flash(f"Erreur : {e}", "danger")
    return render_template('utilisateurs/ajouter.html')

@app.route('/utilisateurs/<int:uid>/desactiver', methods=['POST'])
@login_required
@admin_required
def desactiver_utilisateur(uid):
    if uid == session['user_id']:
        flash("Vous ne pouvez pas désactiver votre propre compte.", "danger")
    else:
        db.update_user(uid, is_active=0)
        flash("Utilisateur désactivé.", "success")
    return redirect(url_for('liste_utilisateurs'))

# -------------------------------------------------------------------
# Paramètres
# -------------------------------------------------------------------
# ============================================================
# Gestion des tolérances (paramètres globaux et individuels)
# ============================================================

@app.route('/parametres/tolerances', methods=['GET', 'POST'])
@login_required
def parametres_tolerances():
    """Paramètres globaux (horaires, tolérances, pénalités)"""
    if request.method == 'POST':
        # Sauvegarde des paramètres
        db.set_parametre('heure_debut_journee', request.form.get('heure_debut', '08:00:00') + ':00')
        db.set_parametre('heure_fin_journee', request.form.get('heure_fin', '17:00:00') + ':00')
        db.set_parametre('duree_pause', request.form.get('duree_pause', '01:00:00') + ':00')
        db.set_parametre('tolerance_retard_globale', request.form.get('tolerance_retard', '10'))
        db.set_parametre('tolerance_depart_anticipe', request.form.get('tolerance_depart', '10'))
        db.set_parametre('seuil_justification_retard', request.form.get('seuil_justif', '15'))
        db.set_parametre('penalite_retard', request.form.get('penalite', '0.50'))
        db.set_parametre('arrondir_retard', '1' if request.form.get('arrondir') == 'on' else '0')
        flash("Paramètres globaux sauvegardés", "success")
        return redirect(url_for('parametres_tolerances'))
    
    # Récupération des valeurs actuelles
    now = datetime.now()
    return render_template('parametres/tolerances.html',
                           heure_debut=db.get_parametre('heure_debut_journee', '08:00:00')[:5],
                           heure_fin=db.get_parametre('heure_fin_journee', '17:00:00')[:5],
                           duree_pause=db.get_parametre('duree_pause', '01:00:00')[:5],
                           tolerance_retard=db.get_parametre('tolerance_retard_globale', '10'),
                           tolerance_depart=db.get_parametre('tolerance_depart_anticipe', '10'),
                           seuil_justif=db.get_parametre('seuil_justification_retard', '15'),
                           penalite=db.get_parametre('penalite_retard', '0.50'),
                           arrondir=db.get_parametre('arrondir_retard', '0') == '1',
                           now=now)

#-------------------------------------------------------------------------------
@app.route('/parametres/tolerances/individuelles')
@login_required
def tolerances_individuelles():
    """Liste des agents avec leurs tolérances personnalisées (pagination)"""
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 20, type=int)
    search = request.args.get('search', '')
    
    if search:
        agents = db.search_personnel(search)
    else:
        agents = db.get_personnel()
    
    total = len(agents)
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    offset = (page - 1) * per_page
    agents_page = agents[offset:offset+per_page]
    
    filter_args = {'search': search} if search else {}
    
    return render_template('parametres/tolerances_individuelles.html',
                           agents=agents_page,
                           page=page,
                           total_pages=total_pages,
                           per_page=per_page,
                           filter_args=filter_args,
                           search=search)

#--------------------------------------------------------------------------------------------------
@app.route('/parametres/tolerances/agent/<int:aid>', methods=['GET', 'POST'])
@login_required
def modifier_tolerances_agent(aid):
    """Modifier les tolérances d'un agent spécifique"""
    agent = db.get_personnel(aid)
    if not agent:
        flash("Agent non trouvé", "danger")
        return redirect(url_for('tolerances_individuelles'))
    
    if request.method == 'POST':
        try:
            heure_entree = request.form.get('heure_entree', '08:00:00') + ':00'
            heure_sortie = request.form.get('heure_sortie', '17:00:00') + ':00'
            tolerance_entree = int(request.form.get('tolerance_entree', 10))
            tolerance_sortie = int(request.form.get('tolerance_sortie', 10))
            type_quart = request.form.get('type_quart', 'jour')
            concerne = 1 if request.form.get('concerne') == 'on' else 0
            
            # Mise à jour des champs spécifiques
            data = (
                agent['matricule'],
                agent.get('badge_id'),
                agent['nom'],
                agent['prenom'],
                agent.get('type_person', 'cadre'),
                agent.get('fonction'),
                agent.get('activite_id'),
                agent.get('division_id'),
                agent.get('direction_id'),
                agent.get('departement_id'),
                agent.get('service_id'),
                agent.get('equipe_id'),
                agent.get('date_embauche'),
                agent.get('date_naissance'),
                agent.get('adresse'),
                agent.get('telephone'),
                agent.get('email'),
                agent.get('photo'),
                agent.get('statut', 'actif'),
                concerne,
                type_quart,
                heure_entree,
                heure_sortie,
                tolerance_entree,
                tolerance_sortie
            )
            db.update_personnel(aid, data)
            flash("Tolérances mises à jour", "success")
            return redirect(url_for('tolerances_individuelles'))
        except Exception as e:
            flash(f"Erreur : {str(e)}", "danger")
    
    return render_template('parametres/tolerances_agent.html', agent=agent)

#---------------------------------------------------------------------------
@app.route('/parametres/tolerances/appliquer_tous', methods=['POST'])
@login_required
def appliquer_tolerances_a_tous():
    """Appliquer les tolérances globales à tous les agents"""
    if request.method == 'POST':
        tol_entree = int(request.form.get('tolerance_entree', 10))
        tol_sortie = int(request.form.get('tolerance_sortie', 10))
        heure_entree = request.form.get('heure_entree', '08:00:00') + ':00'
        heure_sortie = request.form.get('heure_sortie', '17:00:00') + ':00'
        
        with db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE personnel 
                SET heure_entree_theorique = ?, heure_sortie_theorique = ?,
                    tolerance_entree = ?, tolerance_sortie = ?
            """, (heure_entree, heure_sortie, tol_entree, tol_sortie))
            conn.commit()
            flash(f"{cursor.rowcount} agents mis à jour", "success")
        return redirect(url_for('tolerances_individuelles'))

#---------------------------------------------------------------------------------------
@app.route('/parametres/conservation', methods=['GET', 'POST'])
@login_required
@admin_required
def parametres_conservation():
    if request.method == 'POST':
        if 'purge' in request.form:
            resultats = db.purger_anciennes_donnees()
            flash(f"Purge effectuée : {resultats['pointages']} pointages, {resultats['retards']} retards, {resultats['logs']} logs supprimés.", "success")
        else:
            nouvelle_duree = request.form['duree_conservation']
            db.set_parametre('duree_conservation_mois', nouvelle_duree)
            flash(f"Délai de conservation modifié à {nouvelle_duree} mois.", "success")
        return redirect(url_for('parametres_conservation'))
    duree = db.get_parametre('duree_conservation_mois', '12')
    return render_template('parametres/conservation.html', duree=duree)

#------------------------------------------------------------------------------------------------
@app.route('/parametres/weekend', methods=['GET', 'POST'])
@login_required
@admin_required
def parametres_weekend():
    if request.method == 'POST':
        jours = ','.join(request.form.getlist('jours'))
        db.set_parametre('jours_weekend', jours)
        flash("Configuration des jours de weekend enregistrée.", "success")
        return redirect(url_for('parametres_weekend'))
    actuel = db.get_parametre('jours_weekend', '5,6')
    jours_selectionnes = [int(x) for x in actuel.split(',') if x.strip()]
    return render_template('parametres/weekend.html', jours_selectionnes=jours_selectionnes)

# -------------------------------------------------------------------
# API JSON pour badges
# -------------------------------------------------------------------
@app.route('/api/pointage', methods=['POST'])
def api_pointage():
    data = request.json
    identifiant = data.get('matricule') or data.get('badge_id')
    type_pt = data.get('type', 'entrée')
    pers = db.get_personnel_by_matricule(identifiant) or db.get_personnel_by_badge(identifiant)
    if not pers:
        return jsonify({'success': False, 'message': 'Personnel non trouvé'}), 404
    pid, msg = db.add_pointage_avance(
        matricule=pers['matricule'],
        type_pointage=type_pt,
        mode='api',
        user_id=None
    )
    if pid:
        return jsonify({'success': True, 'message': msg})
    else:
        return jsonify({'success': False, 'message': msg}), 400

# -------------------------------------------------------------------
# Paramètres des horaires globaux
# -------------------------------------------------------------------
@app.route('/parametres/horaires_globaux', methods=['GET', 'POST'])
@login_required
@admin_required
def parametres_horaires_globaux():
    if request.method == 'POST':
        db.set_parametre('heure_entree_globale', request.form['heure_entree_globale'])
        db.set_parametre('heure_sortie_globale', request.form['heure_sortie_globale'])
        db.set_parametre('utiliser_horaires_globaux', '1' if 'utiliser_horaires_globaux' in request.form else '0')
        flash("Horaires globaux mis à jour.", "success")
        return redirect(url_for('parametres_horaires_globaux'))
    heure_entree = db.get_parametre('heure_entree_globale', '08:00:00')
    heure_sortie = db.get_parametre('heure_sortie_globale', '17:00:00')
    utiliser_globaux = db.get_parametre('utiliser_horaires_globaux', '0') == '1'
    return render_template('parametres/horaires_globaux.html',
                           heure_entree=heure_entree,
                           heure_sortie=heure_sortie,
                           utiliser_globaux=utiliser_globaux)

# -------------------------------------------------------------------
# Route pour servir les photos des agents
# -------------------------------------------------------------------
@app.route('/photos/<path:filename>')
@login_required
def servir_photo(filename):
    dossier_photos = os.path.join('pointage_data', 'photos')
    return send_from_directory(dossier_photos, filename)

#--------------------------------------------------------------------
# Fonction de pagination générique
#--------------------------------------------------------------------
def paginate(query, page, per_page=20, params=None):
    """Exécute une requête avec pagination.
       Retourne (items, total, page, total_pages)."""
    if params is None:
        params = []
    # Compter le total
    count_query = f"SELECT COUNT(*) FROM ({query}) AS sub"
    with db.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(count_query, params)
        total = cursor.fetchone()[0]
    total_pages = (total + per_page - 1) // per_page if total > 0 else 1
    # Appliquer LIMIT et OFFSET
    offset = (page - 1) * per_page
    paginated_query = query + " LIMIT ? OFFSET ?"
    params_pag = params + [per_page, offset]
    cursor.execute(paginated_query, params_pag)
    items = [dict(row) for row in cursor.fetchall()]
    return items, total, page, total_pages

#----------------------------------------------------------------------
# Export PDF de tous les rapports
#----------------------------------------------------------------------
from weasyprint import HTML
from flask import make_response, jsonify

@app.route('/rapports/export_all', methods=['POST'])
@login_required
def exporter_rapport_pdf_all():
    """Export PDF de toutes les données (non paginées)"""
    try:
        data = request.get_json()
        type_rapport = data.get('type')          # 'presence_resume' ou 'presence_detail'
        date_debut = data.get('date_debut')
        date_fin = data.get('date_fin')
        direction = data.get('direction')
        departement = data.get('departement')
        service = data.get('service')
        equipe = data.get('equipe')
        search = data.get('search')
        colonnes_indices = data.get('colonnes', [])   # indices des colonnes à conserver

        # Récupérer toutes les données (sans pagination)
        if type_rapport == 'presence_resume':
            resultats = db.get_presence_report(
                date_debut, date_fin, search,
                direction, departement, service, equipe,
                type_rapport='resume'
            )
            if not resultats:
                return jsonify({'error': 'Aucune donnée pour cette période'}), 400

            # Mapper les indices vers les noms de colonnes réelles
            colonnes_map = [
                'matricule', 'nom', 'prenom', 'direction',
                'departement', 'service', 'equipe',
                'jours_presence', 'taux_presence'
            ]
            colonnes_conservees = [colonnes_map[i] for i in colonnes_indices if i < len(colonnes_map)]
            df = pd.DataFrame(resultats)[colonnes_conservees]
            titre = f"Rapport de présence - Résumé complet du {date_debut} au {date_fin}"

        elif type_rapport == 'presence_detail':
            details = db.get_presence_report(
                date_debut, date_fin, search,
                direction, departement, service, equipe,
                type_rapport='detail'
            )
            if not details:
                return jsonify({'error': 'Aucune donnée pour cette période'}), 400

            colonnes_map = [
                'date', 'matricule', 'nom', 'prenom',
                'direction', 'present', 'heure_entree', 'heure_sortie'
            ]
            colonnes_conservees = [colonnes_map[i] for i in colonnes_indices if i < len(colonnes_map)]
            df = pd.DataFrame(details)
            # La colonne 'present' contient un booléen ; on la transforme en texte pour l'affichage
            if 'present' in df.columns:
                df['present'] = df['present'].apply(lambda x: 'Oui' if x else 'Non')
            df = df[colonnes_conservees]
            titre = f"Rapport de présence - Détail complet du {date_debut} au {date_fin}"
        
        elif type_rapport == 'absences_resume':
            resultats = db.get_absence_report(
                date_debut=date_debut, date_fin=date_fin, search=search,
                direction_nom=direction, departement_nom=departement,
                service_nom=service, equipe_nom=equipe, type_rapport='resume'
            )
            colonnes_map = ['matricule', 'nom', 'prenom', 'direction', 'departement', 'service', 'equipe', 'jours_ouvres', 'pointes', 'conges', 'absences']
            colonnes_conservees = [colonnes_map[i] for i in colonnes_indices if i < len(colonnes_map)]
            df = pd.DataFrame(resultats)[colonnes_conservees]
            titre = f"Rapport des absences - Résumé complet du {date_debut} au {date_fin}"

        elif type_rapport == 'absences_detail':
            details = db.get_absence_report(
                date_debut=date_debut, date_fin=date_fin, search=search,
                direction_nom=direction, departement_nom=departement,
                service_nom=service, equipe_nom=equipe, type_rapport='detail'
            )
            colonnes_map = ['date', 'matricule', 'nom', 'prenom', 'direction', 'type_absence']
            colonnes_conservees = [colonnes_map[i] for i in colonnes_indices if i < len(colonnes_map)]
            df = pd.DataFrame(details)[colonnes_conservees]
            titre = f"Rapport des absences - Détail complet du {date_debut} au {date_fin}"

        else:
            return jsonify({'error': 'Type de rapport inconnu'}), 400

        if df.empty:
            return jsonify({'error': 'Aucune donnée après sélection des colonnes'}), 400

        # Générer le PDF en paysage
        import tempfile, os
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            pdf_path = tmp.name
        dataframe_to_pdf(df, titre, pdf_path, orientation='landscape')

        with open(pdf_path, 'rb') as f:
            pdf_data = f.read()
        os.unlink(pdf_path)

        response = make_response(pdf_data)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename=rapport_complet.pdf'
        return response

    except Exception as e:
        return jsonify({'error': str(e)}), 500
#----------------------------------------------------------------------
# Lancement de l'import des pointages
#----------------------------------------------------------------------
@app.route('/import/pointages/lancer', methods=['POST'])
@login_required
def lancer_import_pointages():
    filepath = request.json.get('filepath')
    format_colonnes = request.json.get('format_colonnes')  # {'matricule': 'col1', 'datetime': 'col2'}
    if not filepath:
        return jsonify({'error': 'Chemin fichier manquant'}), 400
    
    # Générer un ID de tâche
    import_id = str(int(time() * 1000))
    
    # Créer une file de communication
    q = queue.Queue()
    
    # Lancer l'import dans un thread
    def run_import():
        try:
            # On passe la queue pour envoyer les mises à jour
            result = db.import_pointages_from_file(
                filepath=filepath,
                format_colonnes=format_colonnes,
                user_id=session['user_id'],
                progress_queue=q
            )
            q.put({'status': 'done', 'result': result})
        except Exception as e:
            q.put({'status': 'error', 'message': str(e)})
    
    thread = threading.Thread(target=run_import)
    thread.daemon = True
    thread.start()
    
    import_tasks[import_id] = {'queue': q, 'status': 'running', 'progress': 0, 'total': 0}
    
    return jsonify({'task_id': import_id})
#-----------------------------------------------------------------
from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration

@app.route('/pointages/export_pdf_complet', methods=['POST'])
@login_required
def export_pdf_complet():
    """Exporte le tableau HTML complet (couleurs, colonnes sélectionnées) en PDF"""
    try:
        html_content = request.json.get('html')
        titre = request.json.get('titre', 'Rapport')
        sous_titre = request.json.get('sous_titre', '')
        
        if not html_content:
            return jsonify({'error': 'Aucun contenu HTML'}), 400
        
        # Construction du document HTML complet pour le PDF
        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>{titre}</title>
            <style>
                /* Copiez ici les styles utilisés dans votre template */
                body {{
                    font-family: DejaVu Sans, Arial, sans-serif;
                    font-size: 10pt;
                    margin: 1.5cm;
                }}
                h1 {{
                    color: #2c3e50;
                    text-align: center;
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    margin-top: 20px;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 6px;
                    text-align: left;
                }}
                th {{
                    background-color: #3498db;
                    color: white;
                }}
                .subtotal-row {{
                    background-color: #e9ecef !important;
                    font-weight: bold;
                    border-top: 2px solid #adb5bd;
                }}
                .subtotal-retard {{
                    background-color: #fff3cd !important;  /* jaune pâle */
                }}
                .subtotal-anticipe {{
                    background-color: #d1ecf1 !important;  /* bleu pâle */
                }}
                .text-danger {{
                    color: red !important;
                }}
                .text-warning {{
                    color: orange !important;
                }}
            </style>
        </head>
        <body>
            <h1>{titre}</h1>
            <p>{sous_titre}</p>
            {html_content}
            <p style="text-align: center; margin-top: 30px; font-size: 8pt;">Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </body>
        </html>
        """
        
        font_config = FontConfiguration()
        pdf = HTML(string=full_html).write_pdf(font_config=font_config)
        
        response = make_response(pdf)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename=rapport_retards_complet.pdf'
        return response
        
    except Exception as e:
        print(f"Erreur PDF: {e}")
        return jsonify({'error': str(e)}), 500
#--------------------------------------------------------------------
# Pointage avancé
#--------------------------------------------------------------------
@app.route('/pointages/avance', methods=['GET', 'POST'])
@login_required
def pointage_avance():
    """Pointage avancé avec sélection d'agent et type"""
    agents = db.get_personnel()
    
    # Statistiques système
    try:
        import psutil
        cpu_percent = psutil.cpu_percent(interval=0.5)
        ram = psutil.virtual_memory()
        stats = {
            'cpu_percent': cpu_percent,
            'ram_used': ram.used // (1024**2),
            'ram_total': ram.total // (1024**2),
            'ram_percent': ram.percent
        }
    except ImportError:
        stats = {'cpu_percent': 0, 'ram_used': 0, 'ram_total': 0, 'ram_percent': 0}
    except Exception as e:
        print(f"Erreur stats: {e}")
        stats = {'cpu_percent': 0, 'ram_used': 0, 'ram_total': 0, 'ram_percent': 0}
    
    now = datetime.now()
    
    if request.method == 'POST':
        agent_id = request.form.get('agent_id')
        type_pointage = request.form.get('type_pointage')
        justification = request.form.get('justification', '')
        mode = request.form.get('mode', 'manuel')
        
        if not agent_id or not type_pointage:
            flash("Veuillez sélectionner un agent et un type de pointage", "danger")
            return redirect(url_for('pointage_avance'))
        
        agent = db.get_personnel(agent_id)
        if not agent:
            flash("Agent non trouvé", "danger")
            return redirect(url_for('pointage_avance'))
        
        # Vérifier si l'agent est concerné par le pointage
        if not agent.get('concerne_pointage', 1):
            flash(f"L'agent {agent['prenom']} {agent['nom']} n'est pas concerné par le pointage", "warning")
            return redirect(url_for('pointage_avance'))
        
        # Vérifier les doublons pour le même jour
        today = date.today()
        with db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id FROM pointages 
                WHERE personnel_id = ? AND date_pointage = ? AND type_pointage = ?
            """, (agent['id'], today, type_pointage))
            if cursor.fetchone():
                flash(f"L'agent a déjà un pointage '{type_pointage}' pour aujourd'hui", "warning")
                return redirect(url_for('pointage_avance'))
        
        # Enregistrer le pointage
        now_time = datetime.now()
        quart_id, quart_nom = db.determiner_quart(now_time.strftime('%H:%M:%S'))
        weekend = 1 if today.weekday() >= 5 else 0
        
        with db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO pointages (personnel_id, matricule, badge_id, type_pointage, 
                                       date_pointage, heure_pointage, quart_id, mode, 
                                       est_weekend, justification, user_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (agent['id'], agent['matricule'], agent.get('badge_id'), type_pointage,
                  today, now_time.strftime('%H:%M:%S'), quart_id, mode, weekend, justification,
                  session['user_id']))
            conn.commit()
        
        flash(f"Pointage {type_pointage} enregistré pour {agent['prenom']} {agent['nom']}", "success")
        return redirect(url_for('pointage_avance'))
    
    return render_template('pointages/avance.html', 
                           agents=agents,
                           stats=stats,
                           now=now)

#------------------------------------------------------------------
# 
@app.route('/pointages/import/status/<task_id>')
@login_required
def import_status(task_id):
    """Retourne le statut d'une tâche d'import en cours"""
    if task_id not in import_tasks:
        return jsonify({'error': 'Tâche introuvable'}), 404
    
    task = import_tasks[task_id]
    response = {
        'status': task['status'],
        'progress': task.get('progress', 0),
        'message': task.get('message', ''),
        'total': task.get('total', 0),
        'importes': task.get('results', {}).get('importes', 0) if task.get('results') else 0
    }
    
    # Ajouter les erreurs si la tâche est terminée
    if task.get('results') and task['results'].get('details'):
        response['details'] = task['results']['details'][:10]
    
    return jsonify(response)

#--------------------------------------------------------------------
# Route pour annuler une importation en cours
#--------------------------------------------------------------------
@app.route('/pointages/import/annuler/<task_id>')
@login_required
def import_annuler(task_id):
    """Annule une tâche d'import en cours"""
    if task_id in import_tasks:
        import_tasks[task_id]['status'] = 'cancelled'
        import_tasks[task_id]['message'] = 'Import annulé par l\'utilisateur'
        return jsonify({'success': True})
    return jsonify({'error': 'Tâche introuvable'}), 404

#--------------------------------------------------------------------
# Route pour vérifier le statut de l'import
#--------------------------------------------------------------------
@app.route('/import/pointages/progression/<task_id>')
@login_required
def progression_import(task_id):
    task = import_tasks.get(task_id)
    if not task:
        return jsonify({'status': 'unknown'})
    
    # Récupérer les messages de la queue sans bloquer
    updates = []
    try:
        while True:
            msg = task['queue'].get_nowait()
            updates.append(msg)
            if msg.get('status') in ('done', 'error'):
                # Nettoyer la tâche après un certain temps
                del import_tasks[task_id]
                return jsonify(msg)
            if 'progress' in msg:
                task['progress'] = msg['progress']
                task['total'] = msg.get('total', 0)
    except queue.Empty:
        pass
    
    return jsonify({
        'status': 'running',
        'progress': task['progress'],
        'total': task['total'],
        'updates': updates[-5:]  # derniers messages
    })

#--------------------------------------------------------------------
# Route pour lancer un import de pointages de manière asynchrone
#--------------------------------------------------------------------
@app.route('/pointages/importer_async', methods=['POST'])
@login_required
def importer_pointages_async():
    """Lance un import asynchrone et retourne un task_id"""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Fichier vide'}), 400
    
    # Sauvegarder temporairement le fichier
    temp_dir = os.path.join('pointage_data', 'temp')
    os.makedirs(temp_dir, exist_ok=True)
    filepath = os.path.join(temp_dir, f"import_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}")
    file.save(filepath)
    
    matricule_col = request.form.get('matricule_col', 'matricule')
    datetime_col = request.form.get('datetime_col', 'datetime')
    
    format_colonnes = {'matricule': matricule_col, 'datetime': datetime_col}
    
    # Générer un ID de tâche unique
    task_id = f"task_{datetime.now().strftime('%Y%m%d%H%M%S')}_{os.urandom(4).hex()}"
    
    # Créer une queue pour cette tâche
    import_tasks[task_id] = {
        'status': 'pending',
        'progress': 0,
        'total': 0,
        'message': 'Initialisation...',
        'queue': queue.Queue()
    }
    
    # Lancer le thread d'import
    thread = threading.Thread(
        target=import_task_worker,
        args=(task_id, filepath, format_colonnes, session['user_id'])
    )
    thread.daemon = True
    thread.start()
    
    return jsonify({'task_id': task_id})

#---------------------------------------------------------------------------------
# Export Pointages PDF
#---------------------------------------------------------------------------------

from flask import send_file
import tempfile
import os
from datetime import datetime
import pdfkit  # ou weasyprint

@app.route('/pointages/export_pdf', methods=['POST'])
@login_required
def export_pointages_pdf():
    """Exporte la fiche de pointage en PDF avec entête personnalisée"""
    try:
        # Récupérer les données du formulaire
        date_debut = request.form.get('date_debut', '')
        date_fin = request.form.get('date_fin', '')
        agent_id = request.form.get('agent_id', '')
        matricule = request.form.get('matricule', '')
        nom = request.form.get('nom', '')
        prenom = request.form.get('prenom', '')
        fonction = request.form.get('fonction', '')
        division = request.form.get('division', '')
        departement = request.form.get('departement', '')
        equipe = request.form.get('equipe', '')
        direction = request.form.get('direction', '')
        service = request.form.get('service', '')
        quart = request.form.get('quart', '')
        
        print(f"[DEBUG] Export PDF - date_debut={date_debut}, date_fin={date_fin}, agent_id={agent_id}")
        
        # Récupérer les pointages si une période est spécifiée
        pointages = []
        if date_debut and date_fin:
            if agent_id and agent_id != '':
                # Pointages d'un agent spécifique
                agent = db.get_personnel(int(agent_id))
                if agent:
                    pointages = db.get_pointages_filtres(
                        date_debut=date_debut,
                        date_fin=date_fin,
                        matricule=agent['matricule']
                    )
            else:
                # Pointages de tous les agents
                pointages = db.get_pointages_filtres(
                    date_debut=date_debut,
                    date_fin=date_fin
                )
        
        # Générer le HTML à partir du template
        html = render_template('pointages/export_pdf.html',
                               date_debut=date_debut,
                               date_fin=date_fin,
                               matricule=matricule,
                               nom=nom,
                               prenom=prenom,
                               fonction=fonction,
                               division=division,
                               departement=departement,
                               equipe=equipe,
                               direction=direction,
                               service=service,
                               quart=quart,
                               pointages=pointages,
                               date_generation=datetime.now().strftime('%d/%m/%Y %H:%M'),
                               logo_path=url_for('static', filename='images/logo_entreprise.png', _external=True))
        
        # Convertir HTML en PDF avec WeasyPrint
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
        
        font_config = FontConfiguration()
        
        # Options pour le PDF
        css = CSS(string='''
            @page {
                size: A4;
                margin: 1.5cm;
            }
            body {
                font-family: DejaVu Sans, Arial, sans-serif;
                font-size: 11pt;
            }
        ''', font_config=font_config)
        
        pdf = HTML(string=html).write_pdf(stylesheets=[css], font_config=font_config)
        
        # Envoyer le PDF en réponse
        response = make_response(pdf)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'inline; filename=pointages_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        return response
        
    except Exception as e:
        print(f"[ERROR] Export PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f"Erreur lors de la génération du PDF : {str(e)}", "danger")
        return redirect(request.referrer or url_for('dashboard'))

#-----------------------------------------------------------------------------------------
    
@app.route('/api/system_stats')
@login_required
def api_system_stats():
    """Retourne les statistiques système en temps réel (AJAX)"""
    try:
        import psutil
        cpu_percent = psutil.cpu_percent(interval=0.3)
        ram = psutil.virtual_memory()
        return jsonify({
            'cpu_percent': cpu_percent,
            'ram_used': ram.used // (1024**2),
            'ram_total': ram.total // (1024**2),
            'ram_percent': ram.percent
        })
    except ImportError:
        return jsonify({'cpu_percent': 0, 'ram_used': 0, 'ram_total': 0, 'ram_percent': 0})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

#-------------------------------------------------------------------
# Route pour obtenir des infos système (CPU, RAM) - pour monitoring
#-------------------------------------------------------------------
@app.route('/system/info')
@login_required
def system_info():
    return jsonify({
        'cpu_percent': psutil.cpu_percent(interval=0.5),
        'ram_percent': psutil.virtual_memory().percent
    })
    

# -------------------------------------------------------------------
# Lancement
# -------------------------------------------------------------------
if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5000)


