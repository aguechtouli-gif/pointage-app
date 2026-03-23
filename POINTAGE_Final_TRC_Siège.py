"""
POINTAGE_FINAL.py
Système de Pointage - Version ULTIME
Toutes les classes sont intégrées, plus aucun placeholder.
Auteur: TRC
Date: 2024
"""
from logging import root
import sys
import os

def resource_path(relative_path):
    """Retourne le chemin absolu vers une ressource, fonctionne pour le développement et pour PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def resource_path(relative_path):
    """ Obtient le chemin absolu vers une ressource, fonctionne pour le développement et pour PyInstaller. """
    try:
        # PyInstaller crée un dossier temporaire et stocke le chemin dans _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

icon_path = resource_path('app.ico')


import sqlite3
import os
import hashlib
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import csv
from PIL import Image, ImageTk
import json
import shutil


# ============================================================
#               INTERFACES GRAPHIQUES
# ============================================================

# ------------------------------------------------------------
#   FENÊTRE DE CONNEXION
# ------------------------------------------------------------
class LoginWindow:
    def __init__(self, root, app):
        self.root = root
        self.app = app
        self.root.title("Système de Pointage - Connexion")
        self.root.geometry("400x500")
        self.root.configure(bg='#f0f0f0')
        # Charger l'icône de la fenêtre
        self.charger_icone()
        
        self.create_widgets()
        
    def charger_icone(self):
        """Charge l'icône de la fenêtre à partir du fichier logo."""
        try:
            from PIL import Image, ImageTk
            # Utilisez resource_path si vous l'avez définie
            icon_path = resource_path(os.path.join('assets', 'logo_entreprise.png'))
            if os.path.exists(icon_path):
                img = Image.open(icon_path)
                img = img.resize((32, 32), Image.LANCZOS)
                icon = ImageTk.PhotoImage(img)
                self.root.iconphoto(False, icon)
                # Conserver une référence pour éviter le garbage collector
                self.icon = icon
            else:
                print("Icône non trouvée :", icon_path)
        except Exception as e:
            print(f"Erreur chargement icône : {e}")

    def create_widgets(self):
        tk.Label(self.root, text="SYSTÈME DE POINTAGE", font=('Arial',18,'bold'),
                 bg='#f0f0f0', fg='#2c3e50').pack(pady=30)
        form = tk.Frame(self.root, bg='white', relief=tk.RAISED, bd=2)
        form.pack(pady=20, padx=40, fill=tk.BOTH, expand=True)
        tk.Label(form, text="🔐", font=('Arial',40), bg='white', fg='#3498db').pack(pady=20)
        tk.Label(form, text="Nom d'utilisateur:", font=('Arial',10), bg='white').pack(anchor='w', padx=20, pady=(10,0))
        self.username = tk.Entry(form, font=('Arial',12))
        self.username.pack(fill=tk.X, padx=20, pady=5)
        self.username.focus_set()
        tk.Label(form, text="Mot de passe:", font=('Arial',10), bg='white').pack(anchor='w', padx=20, pady=(10,0))
        self.password = tk.Entry(form, font=('Arial',12), show='*')
        self.password.pack(fill=tk.X, padx=20, pady=5)
        tk.Button(form, text="SE CONNECTER", command=self.login,
                  bg='#3498db', fg='white', font=('Arial',12,'bold'), width=20, height=2).pack(pady=30)
        self.error = tk.Label(form, text="", fg='red', bg='white')
        self.error.pack()
        tk.Button(self.root, text="Quitter", command=self.root.quit, bg='#e74c3c', fg='white').pack(pady=10)
        self.password.bind('<Return>', lambda e: self.login())

    def login(self):
        u = self.username.get().strip()
        p = self.password.get().strip()
        if not u or not p:
            self.error.config(text="Veuillez remplir tous les champs")
            return
        user = self.app.db.authenticate_user(u, p)
        if user:
            self.app.current_user = user
            self.root.destroy()
            self.app.show_main_window()
        else:
            self.error.config(text="Identifiants incorrects")


# ------------------------------------------------------------
#   IMPORTATION DE POINTAGES
# ------------------------------------------------------------
class ImportPointages:
    """Interface d'importation des fichiers de pointages (badgeuse)"""
    
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.window = None
    
    def ouvrir(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Importer des pointages")
        self.window.geometry("750x650")
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.configure(bg='#f0f0f0')
        
        # Titre
        tk.Label(self.window, text="📥 IMPORTATION DE FICHIER DE POINTAGES", 
                font=('Arial', 14, 'bold'), bg='#f0f0f0', fg='#2c3e50').pack(pady=20)
        
        # Frame principal
        main_frame = tk.Frame(self.window, bg='white', relief=tk.RAISED, bd=2)
        main_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        # Instructions
        instr_frame = tk.Frame(main_frame, bg='#e8f4f8')
        instr_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(instr_frame, text="📋 Format attendu :", 
                font=('Arial', 10, 'bold'), bg='#e8f4f8').pack(anchor='w', padx=10, pady=5)
        tk.Label(instr_frame, 
                text="CSV ou Excel contenant : matricule (ou badge), date, heure, type (entrée/sortie/pause_début/pause_fin)",
                bg='#e8f4f8', font=('Arial', 9)).pack(anchor='w', padx=20, pady=2)
        tk.Label(instr_frame, 
                text="Colonnes par défaut : matricule, badge_id, date, heure, type — vous pouvez les mapper manuellement.",
                bg='#e8f4f8', font=('Arial', 9)).pack(anchor='w', padx=20, pady=2)
        
        # Sélection fichier
        file_frame = tk.Frame(main_frame, bg='white')
        file_frame.pack(fill=tk.X, padx=10, pady=15)
        
        tk.Label(file_frame, text="Fichier source :", font=('Arial', 10, 'bold'),
                bg='white').pack(anchor='w', padx=5, pady=5)
        
        path_frame = tk.Frame(file_frame, bg='white')
        path_frame.pack(fill=tk.X)
        
        self.file_path = tk.StringVar()
        tk.Entry(path_frame, textvariable=self.file_path, width=50, 
                font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
        tk.Button(path_frame, text="📂 Parcourir", command=self.parcourir_fichier,
                 bg='#3498db', fg='white', font=('Arial', 10),
                 width=15, height=1).pack(side=tk.LEFT, padx=5)
        
        # Cadre de mapping des colonnes
        mapping_frame = tk.LabelFrame(main_frame, text="Correspondance des colonnes", 
                                     font=('Arial', 10, 'bold'), bg='white')
        mapping_frame.pack(fill=tk.X, padx=10, pady=15)
        
        colonnes_defaut = ['matricule', 'badge_id', 'date', 'heure', 'type']
        self.entries_map = {}
        
        for i, col in enumerate(colonnes_defaut):
            tk.Label(mapping_frame, text=col.capitalize(), bg='white', font=('Arial', 9)).grid(row=i, column=0, padx=5, pady=5, sticky='w')
            self.entries_map[col] = tk.Entry(mapping_frame, width=20)
            self.entries_map[col].grid(row=i, column=1, padx=5, pady=5)
            self.entries_map[col].insert(0, col)  # valeur par défaut
        
        # Options supplémentaires
        opt_frame = tk.Frame(main_frame, bg='white')
        opt_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.ignore_doublons = tk.BooleanVar(value=True)
        tk.Checkbutton(opt_frame, text="Ignorer les doublons (même matricule, date, heure, type)", 
                      variable=self.ignore_doublons, bg='white').pack(anchor='w', pady=2)
        
        # Aperçu
        preview_frame = tk.LabelFrame(main_frame, text="Aperçu du fichier", 
                                     font=('Arial', 10, 'bold'), bg='white')
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.preview_text = tk.Text(preview_frame, height=10, width=70, font=('Courier', 9))
        self.preview_text.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        
        # Boutons
        btn_frame = tk.Frame(self.window, bg='#f0f0f0')
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="✅ IMPORTER", command=self.importer,
                 bg="#f30d11", fg='white', font=('Arial', 12, 'bold'),
                 width=20, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="❌ ANNULER", command=self.fermer,
                 bg="#ea3e2b", fg='white', font=('Arial', 10, 'bold'),
                 width=15, height=1).pack(side=tk.LEFT, padx=10)
    
    def parcourir_fichier(self):
        filename = filedialog.askopenfilename(
            title="Sélectionner le fichier de pointages",
            filetypes=[("Fichiers CSV", "*.csv"), ("Fichiers Excel", "*.xlsx *.xls"), ("Tous", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            self.afficher_apercu(filename)
    
    def afficher_apercu(self, filename):
        self.preview_text.delete(1.0, tk.END)
        try:
            if filename.lower().endswith('.csv'):
                df = pd.read_csv(filename, nrows=5)
            else:
                df = pd.read_excel(filename, nrows=5)
            
            self.preview_text.insert(tk.END, "📋 APERÇU DES 5 PREMIÈRES LIGNES:\n\n")
            self.preview_text.insert(tk.END, df.to_string())
            self.preview_text.insert(tk.END, f"\n\n✅ Fichier: {os.path.basename(filename)}")
            self.preview_text.insert(tk.END, f"\n📊 Colonnes détectées: {', '.join(df.columns)}")
            
            # Remplir automatiquement le mapping avec les colonnes du fichier
            for col_entry in self.entries_map.values():
                col_entry.delete(0, tk.END)
            # On garde les noms par défaut si les colonnes correspondent
            for col in df.columns:
                col_lower = col.lower()
                if 'matricule' in col_lower:
                    self.entries_map['matricule'].insert(0, col)
                elif 'badge' in col_lower or 'badge_id' in col_lower:
                    self.entries_map['badge_id'].insert(0, col)
                elif 'date' in col_lower:
                    self.entries_map['date'].insert(0, col)
                elif 'heure' in col_lower:
                    self.entries_map['heure'].insert(0, col)
                elif 'type' in col_lower or 'sens' in col_lower:
                    self.entries_map['type'].insert(0, col)
            
            # Si une colonne n'a pas été remplie, on laisse le champ vide
            for key, entry in self.entries_map.items():
                if not entry.get():
                    entry.insert(0, key)
                    
        except Exception as e:
            self.preview_text.insert(tk.END, f"❌ Erreur de lecture: {str(e)}")
    
    def importer(self):
        fichier = self.file_path.get().strip()
        if not fichier:
            messagebox.showwarning("Attention", "Veuillez sélectionner un fichier")
            return
        
        # Construire le mapping
        mapping = {}
        for key, entry in self.entries_map.items():
            col_name = entry.get().strip()
            if col_name:
                mapping[key] = col_name
        
        if not messagebox.askyesno("Confirmation", "L'importation des pointages va commencer. Voulez-vous continuer ?"):
            return
        
        try:
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, "⏳ Importation en cours... Veuillez patienter")
            self.window.update()
            
            results = self.app.db.import_pointages_from_file(
                filepath=fichier,
                format_colonnes=mapping,
                user_id=self.app.current_user['id'] if self.app.current_user else None
            )
            
            # Affichage des résultats
            result_text = f"""
📊 RÉSULTATS DE L'IMPORTATION DES POINTAGES
{'='*50}

✅ Fichier: {os.path.basename(fichier)}

📈 STATISTIQUES:
• Total lignes lues: {results['total']}
• ✅ Pointages importés: {results['importes']}
• ⚠️ Doublons ignorés: {results['doublons']}
• ❌ Erreurs: {results['erreurs']}
"""
            if results['details']:
                result_text += f"\n⚠️ DÉTAILS (10 premiers):\n"
                for i, detail in enumerate(results['details'][:10], 1):
                    result_text += f"  {i}. {detail}\n"
                if len(results['details']) > 10:
                    result_text += f"  ... et {len(results['details'])-10} autre(s) détail(s)\n"
            
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, result_text)
            
            messagebox.showinfo("Import terminé", 
                              f"{results['importes']} pointages importés, {results['doublons']} doublons ignorés.")
            
            # Optionnel : fermeture automatique si tout s'est bien passé
            if results['erreurs'] == 0:
                if messagebox.askyesno("Succès", "Tous les pointages ont été importés. Fermer la fenêtre ?"):
                    self.fermer()
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'import: {str(e)}")
    
    def fermer(self):
        if self.window:
            self.window.destroy()

# ------------------------------------------------------------
#   FENÊTRE PRINCIPALE
# ------------------------------------------------------------
class MainWindow:
    def __init__(self, root, app):
        self.root = root
        self.app = app
        self.root.title(f"Système de Pointage - {self.app.current_user['prenom']} {self.app.current_user['nom']}")
        self.root.geometry("1400x800")
        self.root.state('zoomed')
        self.setup_menu()
        self.create_widgets()
        self.show_dashboard()

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Fichier", menu=file_menu)
        file_menu.add_command(label="📊 Tableau de bord", command=self.show_dashboard)
        file_menu.add_separator()
        file_menu.add_command(label="📂 Importer personnel", command=self.importer_personnel)
        file_menu.add_command(label="📤 Exporter données", command=self.exporter_donnees)
        file_menu.add_separator()
        file_menu.add_command(label="🚪 Changer d'utilisateur", command=self.logout)
        file_menu.add_command(label="❌ Quitter", command=self.root.quit)

        gest_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Gestion", menu=gest_menu)
        gest_menu.add_command(label="👥 Personnel", command=self.show_personnel)
        gest_menu.add_command(label="📥 Importer pointages", command=self.importer_pointages)
        gest_menu.add_command(label="🏢 Hiérarchie", command=self.show_hierarchie)
        gest_menu.add_separator()
        gest_menu.add_command(label="⏰ Pointages", command=self.show_pointages)
        gest_menu.add_command(label="💰 Heures supplémentaires", command=self.show_heures_sup)
        gest_menu.add_command(label="📅 Congés", command=self.show_conges)
        gest_menu.add_command(label="🎉 Jours fériés", command=self.show_jours_feries)

        rapp_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Rapports", menu=rapp_menu)
        rapp_menu.add_command(label="📋 Présence", command=self.show_rapport_presence)
        rapp_menu.add_command(label="⏱️ Heures travaillées", command=self.show_rapport_heures)
        rapp_menu.add_command(label="⚠️ Retards cumulés", command=self.show_rapport_retards)
        rapp_menu.add_command(label="💰 Heures supplémentaires", command=self.show_rapport_hs)
        rapp_menu.add_separator()
        rapp_menu.add_command(label="📊 Statistiques", command=self.show_statistiques)

        params_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Paramètres", menu=params_menu)
        params_menu.add_command(label="👤 Mon compte", command=self.show_mon_compte)
        if self.app.current_user['role'] == 'admin':
            params_menu.add_separator()
            params_menu.add_command(label="⚙️ Paramètres généraux", command=self.show_parametres)
            params_menu.add_command(label="🎯 Tolérances", command=self.show_tolerances)
            params_menu.add_command(label="👥 Utilisateurs", command=self.show_users)
            params_menu.add_command(label="📜 Logs système", command=self.show_logs)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Aide", menu=help_menu)
        help_menu.add_command(label="📖 Documentation", command=self.show_help)
        help_menu.add_command(label="ℹ️ À propos", command=self.show_about)

    def create_widgets(self):
        # Frame pour le header
        header_frame = tk.Frame(self.root, bg='#2c3e50', height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # --- Ajout du logo ---
        def __init__(self, root, app):
            self.root = root
            self.app = app
            self.root.title("Système de Pointage - Connexion")
            self.root.geometry("400x500")
            self.root.configure(bg='#f0f0f0')
            self.charger_logo()  # appel de la méthode qui charge le logo
            self.create_widgets()
        try:
            from PIL import Image, ImageTk
            # Chemin vers le logo (ajustez selon l'emplacement réel)
            logo_path = os.path.join('pointage_data', 'assets', 'logo_entreprise.png')
            if os.path.exists(logo_path):
                
                img = Image.open(logo_path)
                img = img.resize((110, 50), Image.LANCZOS)  # redimensionnez selon vos besoins
                self.logo_image = ImageTk.PhotoImage(img)   # conservez une référence dans self
                logo_label = tk.Label(header_frame, image=self.logo_image, bg='#2c3e50')
                logo_label.pack(side=tk.LEFT, padx=10)
            else:
                print("Logo non trouvé :", logo_path)
        except Exception as e:
            print(f"Erreur chargement logo : {e}")

        # Titre (placé après le logo)
        title_label = tk.Label(header_frame, text="SYSTÈME DE POINTAGE",
                           font=('Arial', 20, 'bold'), bg='#2c3e50', fg='white')
        title_label.pack(side=tk.LEFT, padx=20)
        
        
        tk.Label(header_frame, text="SYSTÈME DE POINTAGE", font=('Arial',20,'bold'),
                 bg='#2c3e50', fg='white').pack(side=tk.LEFT, padx=20)
        user_info = tk.Frame(header_frame, bg='#2c3e50')
        user_info.pack(side=tk.RIGHT, padx=20)
        tk.Label(user_info, text=f"{self.app.current_user['prenom']} {self.app.current_user['nom']} ({self.app.current_user['role']})",
                 font=('Arial',12), bg='#2c3e50', fg='white').pack(side=tk.LEFT, padx=10)
        self.content_frame = tk.Frame(self.root, bg='#ecf0f1')
        self.content_frame.pack(fill=tk.BOTH, expand=True)

    def clear_content(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

    # --------------------------------------------------------
    #   TABLEAU DE BORD
    # --------------------------------------------------------
    def show_dashboard(self):
        self.clear_content()
        
    # Ajouter un logo ou une image d'accueil
        from PIL import Image, ImageTk

    # Dans la méthode show_dashboard (ou dans create_widgets)
        logo_path = os.path.join('pointage_data', 'assets', 'logo_entreprise.png')
    
        try:
            img = Image.open(logo_path)
            img = img.resize((500, 150), Image.LANCZOS)  # redimensionner si nécessaire
            logo_image = ImageTk.PhotoImage(img)
            logo_label = tk.Label(self.content_frame, image=logo_image, bg='#ecf0f1')
            logo_label.image = logo_image  # garder une référence
            logo_label.pack(pady=10)
        except Exception as e:
              print(f"Impossible de charger le logo : {e}")      
    
    # Frame principal avec défilement
        canvas = tk.Canvas(self.content_frame, bg='#ecf0f1', highlightthickness=0)
        scrollbar = tk.Scrollbar(self.content_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#ecf0f1')
    
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
       )
    
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
    
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
        # Récupération des statistiques
        stats = self.app.db.get_statistiques_completes()
    
        # ----- Ligne 1 : Cartes de statistiques -----
        cards_frame = tk.Frame(scrollable_frame, bg='#ecf0f1')
        cards_frame.pack(fill=tk.X, padx=20, pady=10)
    
        cards = [
        ("👥 Personnel actif", stats.get('personnel_actif', 0), "#3498db"),
        ("🎯 Concernés pointage", stats.get('personnel_concerne', 0), "#2ecc71"),
        ("⏰ Pointages (mois)", stats.get('pointages_periode', 0), "#e74c3c"),
        ("⚠️ Retards (mois)", f"{stats.get('retards_periode', 0)} ({stats.get('minutes_retard_periode', 0)} min)", "#f39c12"),
        ]
    
        for i, (title, value, color) in enumerate(cards):
            card = tk.Frame(cards_frame, bg=color, relief=tk.RAISED, bd=2, height=100, width=200)
            card.grid(row=0, column=i, padx=10, pady=10, sticky='nsew')
            card.grid_propagate(False)
            cards_frame.columnconfigure(i, weight=1)
        
            tk.Label(card, text=str(value), font=('Arial', 20, 'bold'), 
                bg=color, fg='white').pack(expand=True)
            tk.Label(card, text=title, font=('Arial', 10), 
                bg=color, fg='white').pack()
    
        # ----- Ligne 2 : Graphiques -----
        graph_frame = tk.Frame(scrollable_frame, bg='#ecf0f1')
        graph_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
        # Création de deux sous-graphiques côte à côte
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        from matplotlib.figure import Figure
    
        fig = Figure(figsize=(12, 5), dpi=100)
        fig.subplots_adjust(hspace=0.4, wspace=0.3)
    
        # Graphique 1 : Répartition par direction
        ax1 = fig.add_subplot(121)
        directions = {}
        # On récupère les données de la base (exemple simplifié)
        with self.app.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT direction, COUNT(*) FROM personnel WHERE statut='actif' GROUP BY direction")
            for row in c.fetchall():
                directions[row[0]] = row[1]
    
        if directions:
           ax1.pie(directions.values(), labels=directions.keys(), autopct='%1.1f%%', startangle=90)
           ax1.set_title("Répartition par direction")
        else:
           ax1.text(0.5, 0.5, "Aucune donnée", ha='center', va='center')
           ax1.set_title("Répartition par direction")
    
        # Graphique 2 : Évolution des pointages sur 30 jours
        ax2 = fig.add_subplot(122)
        # Récupérer les pointages des 30 derniers jours
        import datetime
        date_fin = datetime.date.today()
        date_debut = date_fin - datetime.timedelta(days=30)
        with self.app.db.get_connection() as conn:
             c = conn.cursor()
             c.execute("""
                SELECT date_pointage, COUNT(*) 
                FROM pointages 
                WHERE date_pointage BETWEEN ? AND ? 
                GROUP BY date_pointage 
                ORDER BY date_pointage
        """, (date_debut, date_fin))
        dates = []
        counts = []
        for row in c.fetchall():
            dates.append(row[0])
            counts.append(row[1])
        if dates:
           ax2.plot(dates, counts, marker='o', linestyle='-', color='#3498db')
           ax2.set_title("Pointages des 30 derniers jours")
           ax2.set_xlabel("Date")
           ax2.set_ylabel("Nombre")
           ax2.tick_params(axis='x', rotation=45)
        else:
           ax2.text(0.5, 0.5, "Aucune donnée", ha='center', va='center')
           ax2.set_title("Pointages des 30 derniers jours")
    
        # Intégration dans Tkinter
        canvas_graph = FigureCanvasTkAgg(fig, graph_frame)
        canvas_graph.draw()
        canvas_graph.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
        # ----- Ligne 3 : Top 5 des retardataires -----
        top_frame = tk.LabelFrame(scrollable_frame, text="🚨 Top 5 des retardataires (mois en cours)", 
                               font=('Arial', 12, 'bold'), bg='white')
        top_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
    
        columns = ('Matricule', 'Nom', 'Prénom', 'Direction', 'Nb retards', 'Total minutes')
        tree = ttk.Treeview(top_frame, columns=columns, show='headings', height=5)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
            tree.column('Nom', width=120)
            tree.column('Prénom', width=120)
    
        # Récupérer les retards cumulés du mois
        mois = datetime.date.today().strftime('%Y-%m')
        with self.app.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("""
                SELECT r.matricule, p.nom, p.prenom, p.direction,
                   COUNT(r.id) as nb_retards,
                   SUM(r.minutes_retard) as total_min
                FROM retards_cumules r
                JOIN personnel p ON r.personnel_id = p.id
                WHERE r.mois = ?
                GROUP BY r.personnel_id
                ORDER BY total_min DESC
                LIMIT 5
            """, (mois,))
            for row in c.fetchall():
                tree.insert('', 'end', values=(
                    row['matricule'], row['nom'], row['prenom'], row['direction'],
                    row['nb_retards'], f"{row['total_min']} min"
                ))
    
        scrollbar_top = ttk.Scrollbar(top_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar_top.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar_top.pack(side=tk.RIGHT, fill=tk.Y)
    
        # ----- Ligne 4 : Derniers pointages (liste) -----
        last_frame = tk.LabelFrame(scrollable_frame, text="⏱️ Derniers pointages", 
                                font=('Arial', 12, 'bold'), bg='white')
        last_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
    
        columns2 = ('Date', 'Heure', 'Matricule', 'Nom', 'Type', 'Retard')
        tree2 = ttk.Treeview(last_frame, columns=columns2, show='headings', height=8)
        for col in columns2:
            tree2.heading(col, text=col)
            tree2.column(col, width=100)
            tree2.column('Nom', width=150)
    
        derniers = stats.get('derniers_pointages', [])
        for pt in derniers[:15]:
            retard = f"{pt['minutes_retard']} min" if pt.get('minutes_retard', 0) > 0 else "-"
            tree2.insert('', 'end', values=(
                pt['date_pointage'],
                pt['heure_pointage'][:5] if pt['heure_pointage'] else '',
                pt['matricule'],
                f"{pt.get('prenom', '')} {pt.get('nom', '')}",
                pt['type_pointage'],
                retard
            ))
    
        scrollbar_last = ttk.Scrollbar(last_frame, orient=tk.VERTICAL, command=tree2.yview)
        tree2.configure(yscrollcommand=scrollbar_last.set)
        tree2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar_last.pack(side=tk.RIGHT, fill=tk.Y)

        # --------------------------------------------------------
        #   GESTION DU PERSONNEL
        # --------------------------------------------------------
    def show_personnel(self):
        self.clear_content()
        tk.Label(self.content_frame, text="GESTION DU PERSONNEL", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        btn_frame = tk.Frame(self.content_frame, bg='#ecf0f1')
        btn_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Button(btn_frame, text="➕ Ajouter", command=self.add_personnel,
                  bg='#2ecc71', fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="✏️ Modifier", command=self.edit_personnel,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="🗑️ Supprimer", command=self.delete_personnel,
                  bg='#e74c3c', fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="📂 Importer", command=self.importer_personnel,
                  bg='#9b59b6', fg='white').pack(side=tk.LEFT, padx=5)

        search_f = tk.Frame(btn_frame, bg='#ecf0f1')
        search_f.pack(side=tk.RIGHT)
        tk.Label(search_f, text="Recherche:", bg='#ecf0f1').pack(side=tk.LEFT)
        self.search_entry = tk.Entry(search_f, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind('<Return>', lambda e: self.search_personnel())
        tk.Button(search_f, text="🔍", command=self.search_personnel).pack(side=tk.LEFT)

        tree_f = tk.Frame(self.content_frame)
        tree_f.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        cols = ('ID','Matricule','Badge','Photo','Nom','Prénom','Type','Fonction','Direction','Service','Statut')
        self.tree_pers = ttk.Treeview(tree_f, columns=cols, show='headings', height=20)
        wids = [50,100,100,120,120,100,150,120,120,80]
        for c,w in zip(cols,wids):
            self.tree_pers.heading(c, text=c)
            self.tree_pers.column(c, width=w)
        vsb = ttk.Scrollbar(tree_f, orient=tk.VERTICAL, command=self.tree_pers.yview)
        self.tree_pers.configure(yscrollcommand=vsb.set)
        self.tree_pers.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_pers.bind('<Double-Button-1>', lambda e: self.edit_personnel())
        self.charger_personnel()

    def charger_personnel(self):
        for i in self.tree_pers.get_children():
            self.tree_pers.delete(i)
        pers = self.app.db.get_personnel()
        for p in pers:
            self.tree_pers.insert('','end', values=(
                p['id'], p['matricule'], p['badge_id'] or '-',
                p['photo'] or '-', p['nom'], p['prenom'], p['type_person'] or '-',
                p['fonction'], p['direction'], p['service'] or '-', p['statut']
            ))

    def search_personnel(self):
        t = self.search_entry.get().strip()
        if not t:
            self.charger_personnel()
            return
        for i in self.tree_pers.get_children():
            self.tree_pers.delete(i)
        res = self.app.db.search_personnel(t)
        for p in res:
            self.tree_pers.insert('','end', values=(
                p['id'], p['matricule'], p['badge_id'] or '-',
                p['nom'], p['prenom'], p['type_person'] or '-',
                p['fonction'], p['direction'], p['service'] or '-', p['statut']
            ))

    def add_personnel(self):
        dlg = PersonnelDialog(self.root, self.app, None)
        if dlg.result:
            self.charger_personnel()

    def edit_personnel(self):
        sel = self.tree_pers.selection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez un membre du personnel")
            return
        pid = self.tree_pers.item(sel[0])['values'][0]
        dlg = PersonnelDialog(self.root, self.app, pid)
        if dlg.result:
            self.charger_personnel()

    def delete_personnel(self):
        sel = self.tree_pers.selection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez un membre du personnel")
            return
        item = self.tree_pers.item(sel[0])
        pid = item['values'][0]
        nom = item['values'][3]
        prenom = item['values'][4]
        if messagebox.askyesno("Confirmation", f"Supprimer {prenom} {nom} ?"):
            if self.app.db.delete_personnel(pid):
                messagebox.showinfo("Succès", "Personnel supprimé")
                self.charger_personnel()

    # --------------------------------------------------------
    #   GESTION HIÉRARCHIE
    # --------------------------------------------------------
    def show_hierarchie(self):
        gh = GestionHierarchie(self.root, self.app)
        gh.ouvrir()

    # --------------------------------------------------------
    #   IMPORTATION
    # --------------------------------------------------------
    def importer_personnel(self):
        imp = ImportPersonnel(self.root, self.app)
        imp.ouvrir()

    # --------------------------------------------------------
    #   GESTION POINTAGES
    # --------------------------------------------------------
    def show_pointages(self):
        gp = GestionPointages(self.root, self.app)
        gp.ouvrir()

    # --------------------------------------------------------
    #   HEURES SUPPLÉMENTAIRES
    # --------------------------------------------------------
    def show_heures_sup(self):
        gh = GestionHeuresSup(self.root, self.app)
        gh.ouvrir()

    # --------------------------------------------------------
    #   CONGÉS
    # --------------------------------------------------------
    def show_conges(self):
        ga = GestionAbsences(self.root, self.app)
        ga.ouvrir()

    # --------------------------------------------------------
    #   JOURS FÉRIÉS
    # --------------------------------------------------------
    def show_jours_feries(self):
        gj = GestionJoursFeries(self.root, self.app)
        gj.ouvrir()

    # --------------------------------------------------------
    #   RAPPORT PRÉSENCE
    # --------------------------------------------------------
    def show_rapport_presence(self):
        self.clear_content()
        tk.Label(self.content_frame, text="RAPPORT DE PRÉSENCE", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        f_filt = tk.Frame(self.content_frame, bg='#ecf0f1')
        f_filt.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(f_filt, text="Du:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.rp_ddeb = DateEntry(f_filt, width=12, date_pattern='yyyy-mm-dd')
        self.rp_ddeb.pack(side=tk.LEFT, padx=5)
        self.rp_ddeb.set_date(date.today().replace(day=1))
        tk.Label(f_filt, text="Au:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.rp_dfin = DateEntry(f_filt, width=12, date_pattern='yyyy-mm-dd')
        self.rp_dfin.pack(side=tk.LEFT, padx=5)
        self.rp_dfin.set_date(date.today())
        tk.Label(f_filt, text="Direction:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.rp_dir = ttk.Combobox(f_filt, width=20)
        self.rp_dir.pack(side=tk.LEFT, padx=5)
        dirs = self.app.db.get_all_directions()
        self.rp_dir['values'] = ['Toutes'] + [d['nom_direction'] for d in dirs]
        self.rp_dir.set('Toutes')
        tk.Button(f_filt, text="📊 Générer", command=self.generer_rapport_presence,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=10)
        tk.Button(f_filt, text="📤 Exporter", command=self.exporter_rapport_presence,
                  bg='#2ecc71', fg='white').pack(side=tk.LEFT, padx=5)

        tree_f = tk.Frame(self.content_frame)
        tree_f.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        cols = ('Matricule','Nom','Prénom','Fonction','Direction','Jours présence')
        self.tree_rp = ttk.Treeview(tree_f, columns=cols, show='headings', height=20)
        for c in cols:
            self.tree_rp.heading(c, text=c)
            self.tree_rp.column(c, width=120)
        vsb = ttk.Scrollbar(tree_f, orient=tk.VERTICAL, command=self.tree_rp.yview)
        self.tree_rp.configure(yscrollcommand=vsb.set)
        self.tree_rp.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def generer_rapport_presence(self):
        for i in self.tree_rp.get_children():
            self.tree_rp.delete(i)
        ddeb = self.rp_ddeb.get_date().strftime('%Y-%m-%d')
        dfin = self.rp_dfin.get_date().strftime('%Y-%m-%d')
        dir_sel = None if self.rp_dir.get() == 'Toutes' else self.rp_dir.get()
        rapport = self.app.db.get_presence_report(ddeb, dfin, dir_sel)
        for r in rapport:
            self.tree_rp.insert('','end', values=(
                r['matricule'], r['nom'], r['prenom'], r['fonction'], r['direction'], r['jours_presence']
            ))

    def exporter_rapport_presence(self):
        data = []
        for i in self.tree_rp.get_children():
            data.append(self.tree_rp.item(i)['values'])
        if not data:
            messagebox.showwarning("Attention", "Aucune donnée")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".csv",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx")])
        if fn:
            try:
                df = pd.DataFrame(data, columns=['Matricule','Nom','Prénom','Fonction','Direction','Jours présence'])
                if fn.endswith('.csv'):
                    df.to_csv(fn, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(fn, index=False)
                messagebox.showinfo("Succès","Export terminé")
            except Exception as e:
                messagebox.showerror("Erreur", f"Export échoué: {e}")

    # --------------------------------------------------------
    #   RAPPORT HEURES TRAVAILLÉES
    # --------------------------------------------------------
    def show_rapport_heures(self):
        self.clear_content()
        tk.Label(self.content_frame, text="RAPPORT HEURES TRAVAILLÉES", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        f_sel = tk.Frame(self.content_frame, bg='#ecf0f1')
        f_sel.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(f_sel, text="Employé:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.rh_emp = ttk.Combobox(f_sel, width=30)
        self.rh_emp.pack(side=tk.LEFT, padx=5)
        pers = self.app.db.get_personnel()
        self.rh_dict = {f"{p['prenom']} {p['nom']} ({p['matricule']})": p['matricule'] for p in pers if p['statut']=='actif'}
        self.rh_emp['values'] = list(self.rh_dict.keys())
        if self.rh_dict:
            self.rh_emp.set(list(self.rh_dict.keys())[0])
        tk.Label(f_sel, text="Du:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.rh_ddeb = DateEntry(f_sel, width=12, date_pattern='yyyy-mm-dd')
        self.rh_ddeb.pack(side=tk.LEFT, padx=5)
        self.rh_ddeb.set_date(date.today().replace(day=1))
        tk.Label(f_sel, text="Au:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.rh_dfin = DateEntry(f_sel, width=12, date_pattern='yyyy-mm-dd')
        self.rh_dfin.pack(side=tk.LEFT, padx=5)
        self.rh_dfin.set_date(date.today())
        tk.Button(f_sel, text="📈 Afficher", command=self.generer_rapport_heures,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=10)

        res_f = tk.Frame(self.content_frame)
        res_f.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        cols = ('Date', 'Heures', 'Entrée', 'Sortie', 'Pause début', 'Pause fin')
        self.tree_rh = ttk.Treeview(res_f, columns=cols, show='headings', height=20)
        for c in cols:
            self.tree_rh.heading(c, text=c)
            self.tree_rh.column(c, width=120)
        vsb = ttk.Scrollbar(res_f, orient=tk.VERTICAL, command=self.tree_rh.yview)
        self.tree_rh.configure(yscrollcommand=vsb.set)
        self.tree_rh.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def generer_rapport_heures(self):
        for i in self.tree_rh.get_children():
            self.tree_rh.delete(i)
        emp_key = self.rh_emp.get()
        if not emp_key or emp_key not in self.rh_dict:
            messagebox.showwarning("Attention", "Sélectionnez un employé")
            return
        mat = self.rh_dict[emp_key]
        ddeb = self.rh_ddeb.get_date().strftime('%Y-%m-%d')
        dfin = self.rh_dfin.get_date().strftime('%Y-%m-%d')
        hrs = self.app.db.get_heures_travaillees(mat, ddeb, dfin)
        for h in hrs:
            self.tree_rh.insert('','end', values=(h['date'], f"{h['heures']}h", '', '', '', ''))

    # --------------------------------------------------------
    #   RAPPORT RETARDS
    # --------------------------------------------------------
    def show_rapport_retards(self):
        self.clear_content()
        tk.Label(self.content_frame, text="RAPPORT RETARDS CUMULÉS", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        f_filt = tk.Frame(self.content_frame, bg='#ecf0f1')
        f_filt.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(f_filt, text="Mois:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.ret_mois = ttk.Combobox(f_filt, values=['01','02','03','04','05','06','07','08','09','10','11','12'], width=5)
        self.ret_mois.pack(side=tk.LEFT, padx=5)
        self.ret_mois.set(datetime.now().strftime('%m'))
        tk.Label(f_filt, text="Année:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.ret_annee = ttk.Combobox(f_filt, values=[str(y) for y in range(2020,2031)], width=6)
        self.ret_annee.pack(side=tk.LEFT, padx=5)
        self.ret_annee.set(datetime.now().strftime('%Y'))
        tk.Label(f_filt, text="Matricule:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.ret_mat = tk.Entry(f_filt, width=15)
        self.ret_mat.pack(side=tk.LEFT, padx=5)
        tk.Button(f_filt, text="🔍 Rechercher", command=self.generer_rapport_retards,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=10)

        tree_f = tk.Frame(self.content_frame)
        tree_f.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        cols = ('Matricule','Nom','Prénom','Nb retards','Total min','Justifiés','Non justifiés','Mois','Année')
        self.tree_ret = ttk.Treeview(tree_f, columns=cols, show='headings', height=20)
        for c in cols:
            self.tree_ret.heading(c, text=c)
            self.tree_ret.column(c, width=100)
        vsb = ttk.Scrollbar(tree_f, orient=tk.VERTICAL, command=self.tree_ret.yview)
        self.tree_ret.configure(yscrollcommand=vsb.set)
        self.tree_ret.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def generer_rapport_retards(self):
        for i in self.tree_ret.get_children():
            self.tree_ret.delete(i)
        mois = self.ret_mois.get()
        annee = self.ret_annee.get()
        mat = self.ret_mat.get().strip() or None
        retard = self.app.db.get_retards_cumules(mat, f"{annee}-{mois}")
        for r in retard:
            self.tree_ret.insert('','end', values=(
                r['matricule'], r['nom'], r['prenom'],
                r['nombre_retards'], f"{r['total_minutes']} min",
                f"{r['minutes_justifiees']} min",
                f"{r['minutes_non_justifiees']} min",
                mois, annee
            ))

    # --------------------------------------------------------
    #   RAPPORT HEURES SUP
    # --------------------------------------------------------
    def show_rapport_hs(self):
        messagebox.showinfo("Info", "Utilisez le menu Gestion > Heures supplémentaires pour voir et gérer les heures sup.")

    # --------------------------------------------------------
    #   STATISTIQUES
    # --------------------------------------------------------
    def show_statistiques(self):
        self.clear_content()
        tk.Label(self.content_frame, text="STATISTIQUES", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        stats = self.app.db.get_statistiques_completes()
        text = f"""
        📊 Récapitulatif du mois en cours :

        • Personnel actif : {stats['personnel_actif']}
        • Concernés par le pointage : {stats['personnel_concerne']}
        • Pointages enregistrés : {stats['pointages_periode']}
        • Retards : {stats['retards_periode']} ({stats['minutes_retard_periode']} min)

        """
        tk.Label(self.content_frame, text=text, font=('Arial',12), bg='#ecf0f1',
                 justify=tk.LEFT).pack(pady=20, padx=40, anchor='w')

    # --------------------------------------------------------
    #   GESTION DES UTILISATEURS
    # --------------------------------------------------------
    def show_users(self):
        if self.app.current_user['role'] != 'admin':
            messagebox.showwarning("Accès refusé", "Réservé aux administrateurs")
            return
        self.clear_content()
        tk.Label(self.content_frame, text="GESTION DES UTILISATEURS", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        btn_f = tk.Frame(self.content_frame, bg='#ecf0f1')
        btn_f.pack(fill=tk.X, padx=20, pady=10)
        tk.Button(btn_f, text="➕ Ajouter", command=self.add_user,
                  bg='#2ecc71', fg='white').pack(side=tk.LEFT, padx=5)

        tree_f = tk.Frame(self.content_frame)
        tree_f.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        cols = ('ID','Username','Nom','Prénom','Email','Rôle','Actif','Dernière connexion')
        self.tree_users = ttk.Treeview(tree_f, columns=cols, show='headings', height=20)
        for c in cols:
            self.tree_users.heading(c, text=c)
            self.tree_users.column(c, width=120)
        vsb = ttk.Scrollbar(tree_f, orient=tk.VERTICAL, command=self.tree_users.yview)
        self.tree_users.configure(yscrollcommand=vsb.set)
        self.tree_users.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.charger_utilisateurs()

    def charger_utilisateurs(self):
        for i in self.tree_users.get_children():
            self.tree_users.delete(i)
        users = self.app.db.get_all_users()
        for u in users:
            self.tree_users.insert('','end', values=(
                u['id'], u['username'], u['nom'], u['prenom'], u['email'],
                u['role'], 'Oui' if u['is_active'] else 'Non',
                u['last_login'] or 'Jamais'
            ))

    def add_user(self):
        dlg = UserDialog(self.root, self.app)
        if dlg.result:
            self.charger_utilisateurs()

    # --------------------------------------------------------
    #   MON COMPTE
    # --------------------------------------------------------
    def show_mon_compte(self):
        self.clear_content()
        tk.Label(self.content_frame, text="MON COMPTE", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        f = tk.Frame(self.content_frame, bg='white', relief=tk.RAISED, bd=2)
        f.pack(pady=20, padx=50, fill=tk.BOTH, expand=True)
        u = self.app.current_user
        infos = [
            ("Nom d'utilisateur:", u['username']),
            ("Nom:", u['nom']),
            ("Prénom:", u['prenom']),
            ("Email:", u.get('email','')),
            ("Rôle:", u['role']),
            ("Dernière connexion:", u.get('last_login','Jamais')),
        ]
        for i, (lbl, val) in enumerate(infos):
            tk.Label(f, text=lbl, font=('Arial',10,'bold'), bg='white').grid(row=i, column=0, padx=20, pady=5, sticky='w')
            tk.Label(f, text=val, font=('Arial',10), bg='white').grid(row=i, column=1, padx=20, pady=5, sticky='w')
        btn_f = tk.Frame(f, bg='white')
        btn_f.grid(row=len(infos), column=0, columnspan=2, pady=20)
        tk.Button(btn_f, text="✏️ Modifier le profil", command=self.edit_profile,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=10)
        tk.Button(btn_f, text="🔑 Changer mot de passe", command=self.change_password,
                  bg='#2ecc71', fg='white').pack(side=tk.LEFT, padx=10)

    def edit_profile(self):
        dlg = EditProfileDialog(self.root, self.app)
        if dlg.result:
            self.app.current_user = dict(self.app.db.get_user(self.app.current_user['id']))
            self.show_mon_compte()

    def change_password(self):
        dlg = ChangePasswordDialog(self.root, self.app)

    # --------------------------------------------------------
    #   PARAMÈTRES GÉNÉRAUX
    # --------------------------------------------------------
    def show_parametres(self):
        if self.app.current_user['role'] != 'admin':
            messagebox.showwarning("Accès refusé", "Réservé aux administrateurs")
            return
        self.clear_content()
        tk.Label(self.content_frame, text="PARAMÈTRES GÉNÉRAUX", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        f = tk.Frame(self.content_frame, bg='white', relief=tk.RAISED, bd=2)
        f.pack(pady=20, padx=50, fill=tk.BOTH, expand=True)
        params = [
            ('entreprise_nom', 'Nom de l\'entreprise'),
            ('entreprise_adresse', 'Adresse'),
            ('entreprise_telephone', 'Téléphone'),
            ('heure_debut_journee', 'Heure de début de journée'),
            ('heure_fin_journee', 'Heure de fin de journée'),
            ('duree_pause', 'Durée de la pause'),
            ('tolerance_retard_globale', 'Tolérance retard (min)'),
        ]
        self.param_vars = {}
        for i, (cle, lbl) in enumerate(params):
            tk.Label(f, text=lbl+":", font=('Arial',10), bg='white').grid(row=i, column=0, padx=20, pady=5, sticky='w')
            var = tk.StringVar(value=self.app.db.get_parametre(cle, ''))
            self.param_vars[cle] = var
            tk.Entry(f, textvariable=var, width=30).grid(row=i, column=1, padx=20, pady=5)
        tk.Button(f, text="💾 Sauvegarder", command=self.sauvegarder_parametres,
                  bg='#27ae60', fg='white', font=('Arial',11,'bold')).grid(row=len(params), column=0, columnspan=2, pady=20)

    def sauvegarder_parametres(self):
        for cle, var in self.param_vars.items():
            self.app.db.set_parametre(cle, var.get())
        messagebox.showinfo("Succès", "Paramètres sauvegardés")

    # --------------------------------------------------------
    #   TOLÉRANCES
    # --------------------------------------------------------
    def show_tolerances(self):
        gt = GestionTolerances(self.root, self.app)
        gt.ouvrir_interface()

    # --------------------------------------------------------
    #   LOGS
    # --------------------------------------------------------
    def show_logs(self):
        if self.app.current_user['role'] != 'admin':
            messagebox.showwarning("Accès refusé", "Réservé aux administrateurs")
            return
        self.clear_content()
        tk.Label(self.content_frame, text="LOGS SYSTÈME", font=('Arial',16,'bold'),
                 bg='#ecf0f1').pack(pady=10)
        f_filt = tk.Frame(self.content_frame, bg='#ecf0f1')
        f_filt.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(f_filt, text="Niveau:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.log_niv = ttk.Combobox(f_filt, values=['Tous','info','warning','error'], width=10)
        self.log_niv.pack(side=tk.LEFT, padx=5)
        self.log_niv.set('Tous')
        tk.Label(f_filt, text="Catégorie:", bg='#ecf0f1').pack(side=tk.LEFT, padx=5)
        self.log_cat = ttk.Combobox(f_filt, values=['Toutes','authentification','pointage','personnel','systeme','export'], width=15)
        self.log_cat.pack(side=tk.LEFT, padx=5)
        self.log_cat.set('Toutes')
        tk.Button(f_filt, text="🔍 Rechercher", command=self.charger_logs,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=10)

        tree_f = tk.Frame(self.content_frame)
        tree_f.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        cols = ('ID','Date','Niveau','Catégorie','Message','Utilisateur','IP')
        self.tree_logs = ttk.Treeview(tree_f, columns=cols, show='headings', height=20)
        wids = [50,150,80,120,300,120,120]
        for c,w in zip(cols,wids):
            self.tree_logs.heading(c, text=c)
            self.tree_logs.column(c, width=w)
        vsb = ttk.Scrollbar(tree_f, orient=tk.VERTICAL, command=self.tree_logs.yview)
        self.tree_logs.configure(yscrollcommand=vsb.set)
        self.tree_logs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.charger_logs()

    def charger_logs(self):
        for i in self.tree_logs.get_children():
            self.tree_logs.delete(i)
        niv = self.log_niv.get() if hasattr(self,'log_niv') else 'Tous'
        cat = self.log_cat.get() if hasattr(self,'log_cat') else 'Toutes'
        with self.app.db.get_connection() as conn:
            cursor = conn.cursor()
            q = "SELECT l.*, u.username FROM logs l LEFT JOIN utilisateurs u ON l.user_id = u.id WHERE 1=1"
            params = []
            if niv != 'Tous':
                q += " AND l.niveau = ?"
                params.append(niv)
            if cat != 'Toutes':
                q += " AND l.categorie = ?"
                params.append(cat)
            q += " ORDER BY l.created_at DESC LIMIT 500"
            cursor.execute(q, params)
            logs = cursor.fetchall()
            for l in logs:
                self.tree_logs.insert('','end', values=(
                    l['id'], l['created_at'][:19], l['niveau'], l['categorie'],
                    l['message'][:100], l['username'] or 'Système', l['ip_address'] or ''
                ))

    # --------------------------------------------------------
    #   EXPORT / IMPORT
    # --------------------------------------------------------
    def exporter_donnees(self):
        messagebox.showinfo("Info", "Fonctionnalité d'export à développer")

    # --------------------------------------------------------
    #   AIDE
    # --------------------------------------------------------
    def show_help(self):
        help_txt = """
        SYSTÈME DE POINTAGE - AIDE

        📋 FONCTIONNALITÉS:
        • Gestion du personnel avec badge ID et types
        • Hiérarchie (Directions → Activités → Services → Équipes)
        • Système de quarts (Jour/Nuit)
        • Importation de fichiers CSV/Excel
        • Pointage avec détection automatique du quart
        • Calcul des retards avec tolérance
        • Gestion des heures supplémentaires, congés, jours fériés
        • Rapports et statistiques

        👤 COMPTE ADMIN:
        • Identifiant: admin
        • Mot de passe: admin123

        📁 FORMAT D'IMPORTATION:
        • CSV ou Excel avec colonnes: matricule, nom, prenom, direction (obligatoires)

        ⚙️ CONFIGURATION:
        1. Créez d'abord vos directions (Gestion > Hiérarchie)
        2. Ajoutez les activités, services, équipes
        3. Configurez les quarts de travail
        4. Importez ou ajoutez le personnel
        """
        messagebox.showinfo("Aide", help_txt)

    def show_about(self):
        about_txt = """
        SYSTÈME DE POINTAGE
        Version 8.1 - ULTIME

        ✓ Gestion complète du personnel
        ✓ Badge ID et types de personnel
        ✓ Hiérarchie multiniveaux
        ✓ Système de quarts (07h-19h / 19h-07h)
        ✓ Importation de fichiers avec normalisation des matricules
        ✓ Tolérances personnalisables
        ✓ Calcul automatique des retards
        ✓ Gestion des heures sup, congés, jours fériés
        ✓ Rapports complets

        © 2026 - Tous droits réservés
        """
        messagebox.showinfo("À propos", about_txt)

    def logout(self):
        if messagebox.askyesno("Déconnexion", "Voulez-vous vous déconnecter ?"):
            self.root.destroy()
            self.app.show_login()
            
#---------------------------------------------
#    importer_pointages
#---------------------------------------------
    def importer_pointages(self):
        """Importer des pointages depuis un fichier"""
        try:
        # Utiliser la classe définie localement (dans le même fichier)
         importateur = ImportPointages(self.root, self.app)
         importateur.ouvrir()
        except Exception as e:
         messagebox.showerror("Erreur", f"Impossible d'ouvrir l'importateur: {e}")               

# ============================================================
#   CLASSES D'INTERFACE SPÉCIALISÉES
# ============================================================

# ------------------------------------------------------------
#   GESTION POINTAGES
# ------------------------------------------------------------
class GestionPointages:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.window = None
        self.filtres = {}

    def ouvrir(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Gestion des pointages")
        self.window.geometry("1200x700")
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.configure(bg='#f0f0f0')
        self.create_widgets()
        self.charger_pointages()

    def create_widgets(self):
        tk.Label(self.window, text="📋 GESTION DES POINTAGES", 
                font=('Arial',16,'bold'), bg='#f0f0f0', fg='#2c3e50').pack(pady=10)
        main = tk.Frame(self.window, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        quick = tk.LabelFrame(main, text="⏱️ Pointage rapide", font=('Arial',12,'bold'), bg='white')
        quick.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(quick, text="Matricule / Badge:", bg='white').grid(row=0, column=0, padx=5, pady=5)
        self.entry_id = tk.Entry(quick, width=20)
        self.entry_id.grid(row=0, column=1, padx=5, pady=5)
        self.entry_id.focus_set()
        tk.Label(quick, text="Type:", bg='white').grid(row=0, column=2, padx=5, pady=5)
        self.type_var = tk.StringVar(value='entrée')
        ttk.Combobox(quick, textvariable=self.type_var,
                    values=['entrée','sortie','pause_début','pause_fin'], width=12).grid(row=0, column=3, padx=5, pady=5)
        tk.Label(quick, text="Justification:", bg='white').grid(row=0, column=4, padx=5, pady=5)
        self.entry_just = tk.Entry(quick, width=30)
        self.entry_just.grid(row=0, column=5, padx=5, pady=5)
        tk.Button(quick, text="✅ Pointer", command=self.pointer,
                  bg='#27ae60', fg='white', font=('Arial',10,'bold')).grid(row=0, column=6, padx=10, pady=5)

        filt = tk.LabelFrame(main, text="🔍 Filtres", font=('Arial',12,'bold'), bg='white')
        filt.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(filt, text="Du:", bg='white').grid(row=0, column=0, padx=5, pady=5)
        self.date_debut = DateEntry(filt, width=12, date_pattern='yyyy-mm-dd', background='#3498db')
        self.date_debut.grid(row=0, column=1, padx=5, pady=5)
        self.date_debut.set_date(datetime.now().replace(day=1))
        tk.Label(filt, text="Au:", bg='white').grid(row=0, column=2, padx=5, pady=5)
        self.date_fin = DateEntry(filt, width=12, date_pattern='yyyy-mm-dd', background='#3498db')
        self.date_fin.grid(row=0, column=3, padx=5, pady=5)
        self.date_fin.set_date(datetime.now())
        tk.Label(filt, text="Matricule:", bg='white').grid(row=0, column=4, padx=5, pady=5)
        self.filtre_mat = tk.Entry(filt, width=15)
        self.filtre_mat.grid(row=0, column=5, padx=5, pady=5)
        tk.Label(filt, text="Direction:", bg='white').grid(row=1, column=0, padx=5, pady=5)
        self.filtre_dir = ttk.Combobox(filt, width=20)
        self.filtre_dir.grid(row=1, column=1, padx=5, pady=5)
        self.charger_directions()
        tk.Label(filt, text="Type:", bg='white').grid(row=1, column=2, padx=5, pady=5)
        self.filtre_type = ttk.Combobox(filt, values=['','entrée','sortie','pause_début','pause_fin'], width=12)
        self.filtre_type.grid(row=1, column=3, padx=5, pady=5)
        self.filtre_type.set('')
        tk.Button(filt, text="🔍 Rechercher", command=self.appliquer_filtres,
                  bg='#3498db', fg='white').grid(row=1, column=4, columnspan=2, padx=10, pady=5)

        tab = tk.Frame(main, bg='white')
        tab.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID','Date','Heure','Matricule','Nom','Type','Quart','Retard','Justification','Mode')
        self.tree = ttk.Treeview(tab, columns=cols, show='headings', height=15)
        wids = [50,100,80,100,150,100,100,80,200,100]
        for c,w in zip(cols,wids):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w)
        vsb = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tab, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        btnf = tk.Frame(main, bg='white')
        btnf.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btnf, text="📤 Exporter CSV", command=self.exporter_csv,
                  bg='#2ecc71', fg='white').pack(side=tk.RIGHT, padx=5)

    def charger_directions(self):
        try:
            dirs = self.app.db.get_all_directions()
            self.filtre_dir['values'] = [''] + [d['nom_direction'] for d in dirs]
        except:
            pass

    def appliquer_filtres(self):
        self.filtres = {
            'date_debut': self.date_debut.get_date().strftime('%Y-%m-%d'),
            'date_fin': self.date_fin.get_date().strftime('%Y-%m-%d'),
            'matricule': self.filtre_mat.get().strip() or None,
            'type_pointage': self.filtre_type.get().strip() or None,
            'direction': self.filtre_dir.get().strip() or None
        }
        self.charger_pointages()

    def charger_pointages(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        try:
            pts = self.app.db.get_pointages_filtres(**self.filtres)
            for p in pts:
                retard = f"{p['minutes_retard']} min" if p.get('minutes_retard',0) else '-'
                self.tree.insert('','end', values=(
                    p['id'], p['date_pointage'], p['heure_pointage'][:5],
                    p['matricule'], f"{p['prenom']} {p['nom']}",
                    p['type_pointage'], p.get('quart_nom',''),
                    retard, p['justification'] or '', p['mode']
                ))
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger les pointages : {e}")

    def pointer(self):
        ident = self.entry_id.get().strip()
        typ = self.type_var.get()
        just = self.entry_just.get().strip()
        if not ident:
            messagebox.showwarning("Attention", "Saisissez un matricule ou badge")
            return
        pid, msg = self.app.db.add_pointage_avance(
            matricule=ident, badge_id=ident,
            type_pointage=typ, justification=just,
            user_id=self.app.current_user['id']
        )
        if pid:
            messagebox.showinfo("Succès", msg)
            self.entry_id.delete(0, tk.END)
            self.entry_just.delete(0, tk.END)
            self.charger_pointages()
        else:
            messagebox.showerror("Erreur", msg)

    def exporter_csv(self):
        data = []
        for i in self.tree.get_children():
            data.append(self.tree.item(i)['values'])
        if not data:
            messagebox.showwarning("Attention", "Aucune donnée")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".csv",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx")])
        if fn:
            try:
                df = pd.DataFrame(data, columns=['ID','Date','Heure','Matricule','Nom','Type','Quart','Retard','Justif','Mode'])
                if fn.endswith('.csv'):
                    df.to_csv(fn, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(fn, index=False)
                messagebox.showinfo("Succès", "Export terminé")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur export: {e}")


# ------------------------------------------------------------
#   GESTION HEURES SUPPLÉMENTAIRES
# ------------------------------------------------------------
class GestionHeuresSup:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.window = None

    def ouvrir(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Heures supplémentaires")
        self.window.geometry("1100x650")
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.configure(bg='#f0f0f0')
        self.create_widgets()
        self.charger_liste()

    def create_widgets(self):
        tk.Label(self.window, text="💰 HEURES SUPPLÉMENTAIRES", 
                font=('Arial',16,'bold'), bg='#f0f0f0', fg='#2c3e50').pack(pady=10)
        main = tk.Frame(self.window, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        add = tk.LabelFrame(main, text="➕ Nouvelle heure sup", font=('Arial',11,'bold'), bg='white')
        add.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(add, text="Matricule:", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.entry_mat = tk.Entry(add, width=15)
        self.entry_mat.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(add, text="Début (AAAA-MM-JJ HH:MM):", bg='white').grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.entry_deb = tk.Entry(add, width=20)
        self.entry_deb.grid(row=0, column=3, padx=5, pady=5)
        self.entry_deb.insert(0, datetime.now().strftime('%Y-%m-%d %H:%M'))
        tk.Label(add, text="Fin (AAAA-MM-JJ HH:MM):", bg='white').grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.entry_fin = tk.Entry(add, width=20)
        self.entry_fin.grid(row=1, column=3, padx=5, pady=5)
        self.entry_fin.insert(0, (datetime.now()+timedelta(hours=2)).strftime('%Y-%m-%d %H:%M'))
        tk.Label(add, text="Type:", bg='white').grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.type_sup = ttk.Combobox(add, values=['normale','nuit','weekend','jour_ferie'], width=15)
        self.type_sup.set('normale')
        self.type_sup.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(add, text="Motif:", bg='white').grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.entry_motif = tk.Entry(add, width=50)
        self.entry_motif.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky='w')
        tk.Button(add, text="✅ Ajouter", command=self.ajouter,
                  bg='#27ae60', fg='white').grid(row=2, column=4, padx=10, pady=5)

        lst = tk.Frame(main, bg='white')
        lst.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID','Matricule','Nom','Début','Fin','Durée','Type','Statut','Motif')
        self.tree = ttk.Treeview(lst, columns=cols, show='headings', height=15)
        wids = [50,100,150,150,150,80,100,100,200]
        for c,w in zip(cols,wids):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w)
        vsb = ttk.Scrollbar(lst, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        btnf = tk.Frame(main, bg='white')
        btnf.pack(fill=tk.X, padx=10, pady=10)
        if self.app.current_user['role'] in ['admin','superviseur']:
            tk.Button(btnf, text="✅ Approuver", command=self.approuver,
                      bg='#3498db', fg='white').pack(side=tk.LEFT, padx=5)
            tk.Button(btnf, text="❌ Refuser", command=self.refuser,
                      bg='#e74c3c', fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(btnf, text="📤 Exporter", command=self.exporter,
                  bg='#2ecc71', fg='white').pack(side=tk.RIGHT, padx=5)

    def ajouter(self):
        mat = self.entry_mat.get().strip()
        deb = self.entry_deb.get().strip()
        fin = self.entry_fin.get().strip()
        typ = self.type_sup.get()
        motif = self.entry_motif.get().strip()
        if not mat or not deb or not fin:
            messagebox.showwarning("Attention", "Matricule, début et fin obligatoires")
            return
        pers = self.app.db.get_personnel_by_matricule(mat)
        if not pers:
            messagebox.showerror("Erreur", "Matricule inconnu")
            return
        try:
            dt_deb = datetime.strptime(deb, '%Y-%m-%d %H:%M')
            dt_fin = datetime.strptime(fin, '%Y-%m-%d %H:%M')
            duree = (dt_fin - dt_deb).total_seconds() / 3600
            if duree <= 0:
                messagebox.showerror("Erreur", "La date de fin doit être postérieure")
                return
        except:
            messagebox.showerror("Erreur", "Format de date incorrect (AAAA-MM-JJ HH:MM)")
            return
        taux = {'normale':1.25, 'nuit':1.5, 'weekend':1.75, 'jour_ferie':2.0}.get(typ, 1.25)
        hs_id = self.app.db.add_heure_supplementaire(
            pers['id'], mat, deb, fin, duree, typ, taux, motif
        )
        if hs_id:
            messagebox.showinfo("Succès", "Heure supplémentaire ajoutée")
            self.entry_mat.delete(0, tk.END)
            self.entry_motif.delete(0, tk.END)
            self.charger_liste()

    def charger_liste(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        hs_list = self.app.db.get_heures_supplementaires()
        for h in hs_list:
            self.tree.insert('','end', values=(
                h['id'], h['matricule'], f"{h['prenom']} {h['nom']}",
                h['date_heure_debut'], h['date_heure_fin'],
                f"{h['duree_heures']:.1f}h", h['type_heure_sup'],
                h['statut'], h['motif'] or ''
            ))

    def approuver(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez une heure sup")
            return
        hid = self.tree.item(sel[0])['values'][0]
        self.app.db.update_heure_sup_statut(hid, 'approuve', self.app.current_user['username'])
        self.charger_liste()

    def refuser(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez une heure sup")
            return
        hid = self.tree.item(sel[0])['values'][0]
        self.app.db.update_heure_sup_statut(hid, 'refuse', self.app.current_user['username'])
        self.charger_liste()

    def exporter(self):
        data = []
        for i in self.tree.get_children():
            data.append(self.tree.item(i)['values'])
        if not data:
            messagebox.showwarning("Attention", "Aucune donnée")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".csv",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx")])
        if fn:
            try:
                df = pd.DataFrame(data, columns=['ID','Matricule','Nom','Début','Fin','Durée','Type','Statut','Motif'])
                if fn.endswith('.csv'):
                    df.to_csv(fn, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(fn, index=False)
                messagebox.showinfo("Succès", "Export terminé")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur export: {e}")


# ------------------------------------------------------------
#   GESTION JOURS FÉRIÉS
# ------------------------------------------------------------
class GestionJoursFeries:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.window = None

    def ouvrir(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Jours fériés")
        self.window.geometry("600x500")
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.configure(bg='#f0f0f0')
        self.create_widgets()
        self.charger_liste()

    def create_widgets(self):
        tk.Label(self.window, text="🎉 JOURS FÉRIÉS", 
                font=('Arial',16,'bold'), bg='#f0f0f0', fg='#2c3e50').pack(pady=10)
        main = tk.Frame(self.window, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        add = tk.Frame(main, bg='white')
        add.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(add, text="Date:", bg='white').grid(row=0, column=0, padx=5, pady=5)
        self.date_entry = DateEntry(add, width=12, date_pattern='yyyy-mm-dd')
        self.date_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(add, text="Nom:", bg='white').grid(row=0, column=2, padx=5, pady=5)
        self.nom_entry = tk.Entry(add, width=30)
        self.nom_entry.grid(row=0, column=3, padx=5, pady=5)
        tk.Button(add, text="➕ Ajouter", command=self.ajouter,
                  bg='#27ae60', fg='white').grid(row=0, column=4, padx=10, pady=5)

        lst = tk.Frame(main, bg='white')
        lst.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID','Date','Nom','Année')
        self.tree = ttk.Treeview(lst, columns=cols, show='headings', height=15)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120)
        vsb = ttk.Scrollbar(lst, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        btnf = tk.Frame(main, bg='white')
        btnf.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btnf, text="🗑️ Supprimer", command=self.supprimer,
                  bg='#e74c3c', fg='white').pack(side=tk.RIGHT, padx=5)

    def ajouter(self):
        date = self.date_entry.get_date().strftime('%Y-%m-%d')
        nom = self.nom_entry.get().strip()
        if not nom:
            messagebox.showwarning("Attention", "Veuillez saisir un nom")
            return
        self.app.db.add_jour_ferie(date, nom)
        messagebox.showinfo("Succès", "Jour férié ajouté")
        self.nom_entry.delete(0, tk.END)
        self.charger_liste()

    def charger_liste(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        jours = self.app.db.get_all_jours_feries()
        for j in jours:
            self.tree.insert('','end', values=(j['id'], j['date_jour'], j['nom'], j['date_jour'][:4]))

    def supprimer(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez un jour férié")
            return
        jid = self.tree.item(sel[0])['values'][0]
        if messagebox.askyesno("Confirmation", "Supprimer ce jour férié ?"):
            self.app.db.delete_jour_ferie(jid)
            self.charger_liste()


# ------------------------------------------------------------
#   GESTION CONGÉS (ABSENCES) - VERSION COMPLÈTE
# ------------------------------------------------------------
class GestionAbsences:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.window = None
        self.filtres = {}

    def ouvrir(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Gestion des congés")
        self.window.geometry("1300x750")
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.configure(bg='#f0f0f0')
        self.create_widgets()
        self.charger_demandes()
        self.actualiser_statistiques()

    def create_widgets(self):
        tk.Label(self.window, text="📅 GESTION DES CONGÉS", 
                font=('Arial',16,'bold'), bg='#f0f0f0', fg='#2c3e50').pack(pady=10)

        notebook = ttk.Notebook(self.window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_demande = ttk.Frame(notebook)
        notebook.add(self.tab_demande, text="📝 Demande")
        self.create_tab_demande()

        self.tab_liste = ttk.Frame(notebook)
        notebook.add(self.tab_liste, text="📋 Liste")
        self.create_tab_liste()

        if self.app.current_user['role'] in ['admin', 'superviseur']:
            self.tab_approbation = ttk.Frame(notebook)
            notebook.add(self.tab_approbation, text="✅ Approbation")
            self.create_tab_approbation()

        self.tab_stats = ttk.Frame(notebook)
        notebook.add(self.tab_stats, text="📊 Statistiques")
        self.create_tab_stats()

    def create_tab_demande(self):
        main = tk.Frame(self.tab_demande, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        frame_agent = tk.LabelFrame(main, text="Agent concerné", font=('Arial',11,'bold'), bg='white')
        frame_agent.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(frame_agent, text="Matricule :", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.demande_matricule = ttk.Combobox(frame_agent, width=15)
        self.demande_matricule.grid(row=0, column=1, padx=5, pady=5)
        self.charger_liste_matricules()
        tk.Label(frame_agent, text="Nom complet :", bg='white').grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.demande_nom = tk.Entry(frame_agent, width=30, state='readonly')
        self.demande_nom.grid(row=0, column=3, padx=5, pady=5)
        self.demande_matricule.bind('<<ComboboxSelected>>', self.on_agent_selected)

        frame_type = tk.LabelFrame(main, text="Type de congé", font=('Arial',11,'bold'), bg='white')
        frame_type.pack(fill=tk.X, padx=10, pady=10)
        types_conge = ['Congé annuel', 'Congé maladie', 'Congé maternité', 'Congé paternité',
                       'Congé sans solde', 'Congé exceptionnel', 'RTT', 'Formation', 'Événement familial']
        self.type_conge_var = tk.StringVar(value='Congé annuel')
        ttk.Combobox(frame_type, textvariable=self.type_conge_var, values=types_conge,
                     width=30, state='readonly').pack(padx=10, pady=10, anchor='w')

        frame_periode = tk.LabelFrame(main, text="Période", font=('Arial',11,'bold'), bg='white')
        frame_periode.pack(fill=tk.X, padx=10, pady=10)
        per_f = tk.Frame(frame_periode, bg='white')
        per_f.pack(padx=10, pady=10)
        tk.Label(per_f, text="Du :", bg='white').pack(side=tk.LEFT, padx=5)
        self.demande_date_debut = DateEntry(per_f, width=12, date_pattern='yyyy-mm-dd', background='#3498db')
        self.demande_date_debut.pack(side=tk.LEFT, padx=5)
        self.demande_date_debut.set_date(datetime.now())
        tk.Label(per_f, text="au :", bg='white').pack(side=tk.LEFT, padx=5)
        self.demande_date_fin = DateEntry(per_f, width=12, date_pattern='yyyy-mm-dd', background='#3498db')
        self.demande_date_fin.pack(side=tk.LEFT, padx=5)
        self.demande_date_fin.set_date(datetime.now() + timedelta(days=7))
        tk.Label(per_f, text="Durée :", bg='white').pack(side=tk.LEFT, padx=(20,5))
        self.demande_duree = tk.Label(per_f, text="7 jours", font=('Arial',11,'bold'), fg='#27ae60', bg='white')
        self.demande_duree.pack(side=tk.LEFT, padx=5)
        tk.Button(per_f, text="Calculer", command=self.calculer_duree,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=10)

        frame_solde = tk.LabelFrame(main, text="Solde de congés annuels", font=('Arial',11,'bold'), bg='white')
        frame_solde.pack(fill=tk.X, padx=10, pady=10)
        self.solde_label = tk.Label(frame_solde, text="---", font=('Arial',11), fg='#e67e22', bg='white')
        self.solde_label.pack(padx=10, pady=10, anchor='w')

        frame_motif = tk.LabelFrame(main, text="Motif / Remarques", font=('Arial',11,'bold'), bg='white')
        frame_motif.pack(fill=tk.X, padx=10, pady=10)
        self.demande_motif = tk.Text(frame_motif, height=4, width=70)
        self.demande_motif.pack(padx=10, pady=10)

        frame_justif = tk.LabelFrame(main, text="Justificatif (optionnel)", font=('Arial',11,'bold'), bg='white')
        frame_justif.pack(fill=tk.X, padx=10, pady=10)
        jf = tk.Frame(frame_justif, bg='white')
        jf.pack(padx=10, pady=10, fill=tk.X)
        self.justificatif_path = tk.StringVar()
        tk.Entry(jf, textvariable=self.justificatif_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(jf, text="📎 Parcourir", command=self.parcourir_justificatif,
                  bg='#95a5a6', fg='white').pack(side=tk.LEFT, padx=5)

        btn_frame = tk.Frame(main, bg='white')
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="📨 SOUMETTRE LA DEMANDE", command=self.soumettre_demande,
                  bg='#27ae60', fg='white', font=('Arial',12,'bold'),
                  width=30, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="🗑️ EFFACER", command=self.effacer_formulaire,
                  bg='#e74c3c', fg='white', font=('Arial',10),
                  width=15, height=1).pack(side=tk.LEFT, padx=10)

    def create_tab_liste(self):
        main = tk.Frame(self.tab_liste, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        filt_frame = tk.Frame(main, bg='white')
        filt_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(filt_frame, text="Période du :", bg='white').grid(row=0, column=0, padx=5, pady=5)
        self.filtre_date_debut = DateEntry(filt_frame, width=12, date_pattern='yyyy-mm-dd')
        self.filtre_date_debut.grid(row=0, column=1, padx=5, pady=5)
        self.filtre_date_debut.set_date(date.today().replace(day=1))
        tk.Label(filt_frame, text="au :", bg='white').grid(row=0, column=2, padx=5, pady=5)
        self.filtre_date_fin = DateEntry(filt_frame, width=12, date_pattern='yyyy-mm-dd')
        self.filtre_date_fin.grid(row=0, column=3, padx=5, pady=5)
        self.filtre_date_fin.set_date(date.today())
        tk.Label(filt_frame, text="Statut :", bg='white').grid(row=0, column=4, padx=5, pady=5)
        self.filtre_statut = ttk.Combobox(filt_frame, values=['Tous', 'en_attente', 'approuve', 'refuse'], width=15)
        self.filtre_statut.grid(row=0, column=5, padx=5, pady=5)
        self.filtre_statut.set('Tous')
        tk.Label(filt_frame, text="Matricule :", bg='white').grid(row=1, column=0, padx=5, pady=5)
        self.filtre_matricule = tk.Entry(filt_frame, width=15)
        self.filtre_matricule.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(filt_frame, text="Type :", bg='white').grid(row=1, column=2, padx=5, pady=5)
        self.filtre_type = ttk.Combobox(filt_frame, values=['Tous', 'Congé annuel', 'Congé maladie',
                                                             'Congé maternité', 'Congé paternité',
                                                             'Congé sans solde', 'Congé exceptionnel',
                                                             'RTT', 'Formation', 'Événement familial'], width=20)
        self.filtre_type.grid(row=1, column=3, padx=5, pady=5)
        self.filtre_type.set('Tous')
        tk.Button(filt_frame, text="🔍 Filtrer", command=self.appliquer_filtres,
                  bg='#3498db', fg='white').grid(row=1, column=4, padx=10, pady=5)
        tk.Button(filt_frame, text="🔄 Réinitialiser", command=self.charger_demandes,
                  bg='#95a5a6', fg='white').grid(row=1, column=5, padx=5, pady=5)

        tree_frame = tk.Frame(main, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID', 'Matricule', 'Agent', 'Type', 'Début', 'Fin', 'Durée', 'Statut', 'Motif', 'Date demande')
        self.tree_liste = ttk.Treeview(tree_frame, columns=cols, show='headings', height=15)
        col_widths = [50, 100, 150, 150, 100, 100, 70, 100, 200, 120]
        for c, w in zip(cols, col_widths):
            self.tree_liste.heading(c, text=c)
            self.tree_liste.column(c, width=w)
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_liste.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree_liste.xview)
        self.tree_liste.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree_liste.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_liste.bind('<Double-Button-1>', self.afficher_details_demande)

        btn_export = tk.Frame(main, bg='white')
        btn_export.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btn_export, text="📤 Exporter CSV", command=self.exporter_liste,
                  bg='#2ecc71', fg='white').pack(side=tk.RIGHT, padx=5)

    def create_tab_approbation(self):
        main = tk.Frame(self.tab_approbation, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        tk.Label(main, text="Demandes en attente d'approbation",
                font=('Arial',12,'bold'), bg='white', fg='#2c3e50').pack(pady=10)
        tree_frame = tk.Frame(main, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID', 'Matricule', 'Agent', 'Type', 'Début', 'Fin', 'Durée', 'Motif', 'Date demande')
        self.tree_attente = ttk.Treeview(tree_frame, columns=cols, show='headings', height=12)
        col_widths = [50, 100, 150, 150, 100, 100, 70, 200, 120]
        for c, w in zip(cols, col_widths):
            self.tree_attente.heading(c, text=c)
            self.tree_attente.column(c, width=w)
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_attente.yview)
        self.tree_attente.configure(yscrollcommand=vsb.set)
        self.tree_attente.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        btn_frame = tk.Frame(main, bg='white')
        btn_frame.pack(fill=tk.X, padx=10, pady=20)
        tk.Button(btn_frame, text="✅ Approuver", command=self.approuver_demande,
                  bg='#27ae60', fg='white', font=('Arial',11,'bold'),
                  width=15, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="❌ Refuser", command=self.refuser_demande,
                  bg='#e74c3c', fg='white', font=('Arial',11,'bold'),
                  width=15, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="📄 Détails", command=self.voir_details_attente,
                  bg='#3498db', fg='white', width=12).pack(side=tk.LEFT, padx=10)
        self.charger_demandes_attente()

    def create_tab_stats(self):
        main = tk.Frame(self.tab_stats, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        cards_frame = tk.Frame(main, bg='white')
        cards_frame.pack(pady=10)
        self.stat_total = tk.StringVar(value='0')
        self.stat_attente = tk.StringVar(value='0')
        self.stat_approuve = tk.StringVar(value='0')
        self.stat_refuse = tk.StringVar(value='0')
        self.creer_carte(cards_frame, "📊 Total demandes", self.stat_total, "#3498db", 0)
        self.creer_carte(cards_frame, "⏳ En attente", self.stat_attente, "#f39c12", 1)
        self.creer_carte(cards_frame, "✅ Approuvées", self.stat_approuve, "#27ae60", 2)
        self.creer_carte(cards_frame, "❌ Refusées", self.stat_refuse, "#e74c3c", 3)
        graph_frame = tk.Frame(main, bg='white')
        graph_frame.pack(fill=tk.BOTH, expand=True, pady=20)
        self.figure = plt.Figure(figsize=(10,4), dpi=100)
        self.figure.subplots_adjust(hspace=0.4)
        self.ax1 = self.figure.add_subplot(121)
        self.ax2 = self.figure.add_subplot(122)
        self.canvas = FigureCanvasTkAgg(self.figure, graph_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        tk.Button(main, text="🔄 Actualiser", command=self.actualiser_statistiques,
                 bg='#3498db', fg='white').pack(pady=10)

    def creer_carte(self, parent, titre, variable, couleur, col):
        card = tk.Frame(parent, bg=couleur, relief=tk.RAISED, bd=2, width=180, height=90)
        card.grid(row=0, column=col, padx=10, pady=10)
        card.grid_propagate(False)
        tk.Label(card, text=titre, font=('Arial',10), bg=couleur, fg='white').pack(pady=(10,0))
        tk.Label(card, textvariable=variable, font=('Arial',20,'bold'), bg=couleur, fg='white').pack()

    def charger_liste_matricules(self):
        try:
            pers = self.app.db.get_personnel()
            matricules = [p['matricule'] for p in pers if p['statut'] == 'actif']
            self.demande_matricule['values'] = matricules
        except Exception as e:
            print(f"Erreur chargement matricules: {e}")

    def on_agent_selected(self, event=None):
        mat = self.demande_matricule.get()
        if mat:
            pers = self.app.db.get_personnel_by_matricule(mat)
            if pers:
                self.demande_nom.config(state='normal')
                self.demande_nom.delete(0, tk.END)
                self.demande_nom.insert(0, f"{pers['prenom']} {pers['nom']}")
                self.demande_nom.config(state='readonly')
                self.calculer_solde(mat)

    def calculer_solde(self, matricule):
        try:
            quota = 25
            annee = datetime.now().year
            with self.app.db.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT SUM(duree_jours) as total
                    FROM conges
                    WHERE matricule = ? AND type_conge = 'Congé annuel'
                    AND statut IN ('approuve', 'en_attente')
                    AND strftime('%Y', date_debut) = ?
                ''', (matricule, str(annee)))
                row = cursor.fetchone()
                pris = row[0] if row and row[0] else 0
                reste = quota - pris
                self.solde_label.config(text=f"{reste} jours restants sur {quota} jours (année {annee})")
        except Exception as e:
            self.solde_label.config(text="Erreur de calcul du solde")
            print(f"Erreur solde: {e}")

    def calculer_duree(self):
        try:
            debut = self.demande_date_debut.get_date()
            fin = self.demande_date_fin.get_date()
            duree = (fin - debut).days + 1
            if duree > 0:
                self.demande_duree.config(text=f"{duree} jours")
            else:
                self.demande_duree.config(text="0 jour")
                messagebox.showwarning("Attention", "La date de fin doit être postérieure ou égale à la date de début")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur de calcul: {e}")

    def parcourir_justificatif(self):
        filename = filedialog.askopenfilename(
            title="Sélectionner un justificatif",
            filetypes=[("PDF files", "*.pdf"), ("Images", "*.jpg *.png"), ("Tous", "*.*")]
        )
        if filename:
            self.justificatif_path.set(filename)

    def soumettre_demande(self):
        mat = self.demande_matricule.get().strip()
        if not mat:
            messagebox.showwarning("Validation", "Veuillez sélectionner un agent")
            return
        type_conge = self.type_conge_var.get()
        if not type_conge:
            messagebox.showwarning("Validation", "Veuillez sélectionner un type de congé")
            return
        debut = self.demande_date_debut.get_date()
        fin = self.demande_date_fin.get_date()
        if debut > fin:
            messagebox.showwarning("Validation", "La date de fin doit être postérieure ou égale à la date de début")
            return
        duree = (fin - debut).days + 1
        if duree <= 0:
            messagebox.showwarning("Validation", "La durée doit être positive")
            return
        motif = self.demande_motif.get("1.0", tk.END).strip()
        justif = self.justificatif_path.get().strip()
        pers = self.app.db.get_personnel_by_matricule(mat)
        if not pers:
            messagebox.showerror("Erreur", "Matricule inconnu")
            return
        data = (
            pers['id'],
            mat,
            type_conge,
            debut.strftime('%Y-%m-%d'),
            fin.strftime('%Y-%m-%d'),
            duree,
            motif,
            justif,
            'en_attente',
            date.today().strftime('%Y-%m-%d')
        )
        try:
            self.app.db.add_conge(data)
            messagebox.showinfo("Succès", "Demande de congé soumise avec succès")
            self.effacer_formulaire()
            self.charger_demandes()
            self.charger_demandes_attente()
            self.actualiser_statistiques()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'enregistrement: {e}")

    def effacer_formulaire(self):
        self.demande_matricule.set('')
        self.demande_nom.config(state='normal')
        self.demande_nom.delete(0, tk.END)
        self.demande_nom.config(state='readonly')
        self.type_conge_var.set('Congé annuel')
        self.demande_date_debut.set_date(datetime.now())
        self.demande_date_fin.set_date(datetime.now() + timedelta(days=7))
        self.demande_duree.config(text="7 jours")
        self.demande_motif.delete("1.0", tk.END)
        self.justificatif_path.set("")
        self.solde_label.config(text="---")

    def charger_demandes(self):
        for i in self.tree_liste.get_children():
            self.tree_liste.delete(i)
        try:
            demandes = self.app.db.get_conges(self.filtres)
            for d in demandes:
                self.tree_liste.insert('', 'end', values=(
                    d['id'],
                    d['matricule'],
                    f"{d['prenom']} {d['nom']}",
                    d['type_conge'],
                    d['date_debut'],
                    d['date_fin'],
                    d['duree_jours'],
                    self._format_statut(d['statut']),
                    d['motif'][:50] + '...' if d['motif'] and len(d['motif']) > 50 else d['motif'],
                    d.get('date_demande', d['created_at'])[:10]
                ))
        except Exception as e:
            print(f"Erreur chargement demandes: {e}")

    def charger_demandes_attente(self):
        if not hasattr(self, 'tree_attente'):
            return
        for i in self.tree_attente.get_children():
            self.tree_attente.delete(i)
        try:
            filtres = {'statut': 'en_attente'}
            demandes = self.app.db.get_conges(filtres)
            for d in demandes:
                self.tree_attente.insert('', 'end', values=(
                    d['id'],
                    d['matricule'],
                    f"{d['prenom']} {d['nom']}",
                    d['type_conge'],
                    d['date_debut'],
                    d['date_fin'],
                    d['duree_jours'],
                    d['motif'][:50] + '...' if d['motif'] and len(d['motif']) > 50 else d['motif'],
                    d.get('date_demande', d['created_at'])[:10]
                ))
        except Exception as e:
            print(f"Erreur chargement demandes attente: {e}")

    def appliquer_filtres(self):
        self.filtres = {}
        date_debut = self.filtre_date_debut.get_date().strftime('%Y-%m-%d')
        date_fin = self.filtre_date_fin.get_date().strftime('%Y-%m-%d')
        self.filtres['date_debut'] = date_debut
        self.filtres['date_fin'] = date_fin
        statut = self.filtre_statut.get()
        if statut != 'Tous':
            self.filtres['statut'] = statut
        mat = self.filtre_matricule.get().strip()
        if mat:
            self.filtres['matricule'] = mat
        typ = self.filtre_type.get()
        if typ != 'Tous':
            self.filtres['type_conge'] = typ
        self.charger_demandes()

    def approuver_demande(self):
        sel = self.tree_attente.selection()
        if not sel:
            messagebox.showwarning("Attention", "Veuillez sélectionner une demande")
            return
        item = self.tree_attente.item(sel[0])
        demande_id = item['values'][0]
        if messagebox.askyesno("Confirmation", "Approuver cette demande de congé ?"):
            if self.app.db.update_conge_statut(demande_id, 'approuve', self.app.current_user['username']):
                messagebox.showinfo("Succès", "Demande approuvée")
                self.charger_demandes_attente()
                self.charger_demandes()
                self.actualiser_statistiques()
            else:
                messagebox.showerror("Erreur", "Échec de l'approbation")

    def refuser_demande(self):
        sel = self.tree_attente.selection()
        if not sel:
            messagebox.showwarning("Attention", "Veuillez sélectionner une demande")
            return
        item = self.tree_attente.item(sel[0])
        demande_id = item['values'][0]
        motif_refus = tk.simpledialog.askstring("Motif du refus", "Veuillez indiquer le motif du refus :")
        if motif_refus is not None:
            if self.app.db.update_conge_statut(demande_id, 'refuse', self.app.current_user['username']):
                messagebox.showinfo("Succès", "Demande refusée")
                self.charger_demandes_attente()
                self.charger_demandes()
                self.actualiser_statistiques()
            else:
                messagebox.showerror("Erreur", "Échec du refus")

    def afficher_details_demande(self, event=None):
        sel = self.tree_liste.selection()
        if not sel:
            return
        item = self.tree_liste.item(sel[0])
        demande_id = item['values'][0]
        self._afficher_details(demande_id)

    def voir_details_attente(self):
        sel = self.tree_attente.selection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez une demande")
            return
        item = self.tree_attente.item(sel[0])
        demande_id = item['values'][0]
        self._afficher_details(demande_id)

    def _afficher_details(self, demande_id):
        try:
            with self.app.db.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT c.*, p.nom, p.prenom, p.fonction, p.direction, p.service
                    FROM conges c
                    JOIN personnel p ON c.personnel_id = p.id
                    WHERE c.id = ?
                ''', (demande_id,))
                d = cursor.fetchone()
                if not d:
                    messagebox.showerror("Erreur", "Demande introuvable")
                    return
            details = f"""
            📋 DÉTAILS DE LA DEMANDE DE CONGÉ

            👤 AGENT :
            • Matricule : {d['matricule']}
            • Nom : {d['prenom']} {d['nom']}
            • Fonction : {d['fonction']}
            • Direction : {d['direction']}
            • Service : {d['service'] or '-'}

            📅 PÉRIODE :
            • Type de congé : {d['type_conge']}
            • Du : {d['date_debut']} au {d['date_fin']}
            • Durée : {d['duree_jours']} jours

            📝 MOTIF :
            {d['motif'] or 'Aucun motif'}

            📎 JUSTIFICATIF :
            {d['justificatif'] or 'Aucun justificatif'}

            📊 STATUT :
            • Statut : {self._format_statut(d['statut'])}
            • Date de la demande : {d.get('date_demande', d['created_at'])[:10]}
            • Approuvé par : {d['approuve_par'] or '-'}
            • Date d'approbation : {d['date_approbation'][:10] if d['date_approbation'] else '-'}
            """
            messagebox.showinfo("Détails de la demande", details)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'afficher les détails: {e}")

    def actualiser_statistiques(self):
        try:
            with self.app.db.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM conges")
                total = cursor.fetchone()[0] or 0
                self.stat_total.set(str(total))
                cursor.execute("SELECT COUNT(*) FROM conges WHERE statut='en_attente'")
                attente = cursor.fetchone()[0] or 0
                self.stat_attente.set(str(attente))
                cursor.execute("SELECT COUNT(*) FROM conges WHERE statut='approuve'")
                approuve = cursor.fetchone()[0] or 0
                self.stat_approuve.set(str(approuve))
                cursor.execute("SELECT COUNT(*) FROM conges WHERE statut='refuse'")
                refuse = cursor.fetchone()[0] or 0
                self.stat_refuse.set(str(refuse))

                cursor.execute('''
                    SELECT type_conge, COUNT(*) as cnt
                    FROM conges
                    GROUP BY type_conge
                    ORDER BY cnt DESC
                    LIMIT 6
                ''')
                data_type = cursor.fetchall()
                self.ax1.clear()
                if data_type:
                    types = [d['type_conge'] for d in data_type]
                    counts = [d['cnt'] for d in data_type]
                    self.ax1.bar(types, counts, color='#3498db')
                    self.ax1.set_title('Par type de congé')
                    self.ax1.tick_params(axis='x', rotation=45)
                else:
                    self.ax1.text(0.5, 0.5, 'Aucune donnée', ha='center', va='center')
                    self.ax1.set_title('Par type de congé')

                self.ax2.clear()
                labels = ['En attente', 'Approuvées', 'Refusées']
                valeurs = [attente, approuve, refuse]
                couleurs = ['#f39c12', '#27ae60', '#e74c3c']
                if sum(valeurs) > 0:
                    self.ax2.pie(valeurs, labels=labels, colors=couleurs, autopct='%1.1f%%')
                    self.ax2.set_title('Répartition par statut')
                else:
                    self.ax2.text(0.5, 0.5, 'Aucune donnée', ha='center', va='center')
                    self.ax2.set_title('Répartition par statut')
                self.canvas.draw()
        except Exception as e:
            print(f"Erreur stats congés: {e}")

    def exporter_liste(self):
        data = []
        for i in self.tree_liste.get_children():
            data.append(self.tree_liste.item(i)['values'])
        if not data:
            messagebox.showwarning("Attention", "Aucune donnée à exporter")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".csv",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx")])
        if fn:
            try:
                df = pd.DataFrame(data, columns=['ID','Matricule','Agent','Type','Début','Fin','Durée','Statut','Motif','Date demande'])
                if fn.endswith('.csv'):
                    df.to_csv(fn, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(fn, index=False)
                messagebox.showinfo("Succès", f"Export terminé vers {fn}")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur export: {e}")

    @staticmethod
    def _format_statut(statut):
        if statut == 'en_attente':
            return '⏳ En attente'
        elif statut == 'approuve':
            return '✅ Approuvé'
        elif statut == 'refuse':
            return '❌ Refusé'
        return statut


# ------------------------------------------------------------
#   GESTION HIÉRARCHIE (Directions, Activités, Services, Équipes)
# ------------------------------------------------------------
class GestionHierarchie:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.window = None

    def ouvrir(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Gestion de la hiérarchie")
        self.window.geometry("1000x700")
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.configure(bg='#f0f0f0')

        notebook = ttk.Notebook(self.window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_dir = ttk.Frame(notebook)
        notebook.add(self.tab_dir, text="🏢 Directions")
        self.create_tab_directions()

        self.tab_act = ttk.Frame(notebook)
        notebook.add(self.tab_act, text="📋 Activités")
        self.create_tab_activites()

        self.tab_serv = ttk.Frame(notebook)
        notebook.add(self.tab_serv, text="🔧 Services")
        self.create_tab_services()

        self.tab_eq = ttk.Frame(notebook)
        notebook.add(self.tab_eq, text="👥 Équipes")
        self.create_tab_equipes()

        self.charger_directions()
        self.charger_activites()
        self.charger_services()
        self.charger_equipes()

    # ---------- DIRECTIONS ----------
    def create_tab_directions(self):
        main = tk.Frame(self.tab_dir, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        add_frame = tk.LabelFrame(main, text="Ajouter une direction", font=('Arial',11,'bold'), bg='white')
        add_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(add_frame, text="Nom :", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.dir_nom = tk.Entry(add_frame, width=25)
        self.dir_nom.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(add_frame, text="Code :", bg='white').grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.dir_code = tk.Entry(add_frame, width=15)
        self.dir_code.grid(row=0, column=3, padx=5, pady=5)
        tk.Label(add_frame, text="Responsable :", bg='white').grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.dir_resp = tk.Entry(add_frame, width=20)
        self.dir_resp.grid(row=0, column=5, padx=5, pady=5)
        tk.Button(add_frame, text="➕ Ajouter", command=self.ajouter_direction,
                  bg='#27ae60', fg='white').grid(row=0, column=6, padx=10, pady=5)

        list_frame = tk.Frame(main, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID', 'Nom', 'Code', 'Responsable', 'Description', 'Active')
        self.tree_dir = ttk.Treeview(list_frame, columns=cols, show='headings', height=12)
        for c in cols:
            self.tree_dir.heading(c, text=c)
            self.tree_dir.column(c, width=120)
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_dir.yview)
        self.tree_dir.configure(yscrollcommand=vsb.set)
        self.tree_dir.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = tk.Frame(main, bg='white')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btn_frame, text="🗑️ Désactiver", command=self.desactiver_direction,
                  bg='#e74c3c', fg='white').pack(side=tk.RIGHT, padx=5)

    def ajouter_direction(self):
        nom = self.dir_nom.get().strip()
        if not nom:
            messagebox.showwarning("Attention", "Le nom est obligatoire")
            return
        self.app.db.add_direction(nom, self.dir_code.get() or None, self.dir_resp.get() or None)
        messagebox.showinfo("Succès", f"Direction '{nom}' ajoutée")
        self.dir_nom.delete(0, tk.END)
        self.dir_code.delete(0, tk.END)
        self.dir_resp.delete(0, tk.END)
        self.charger_directions()

    def charger_directions(self):
        for i in self.tree_dir.get_children():
            self.tree_dir.delete(i)
        dirs = self.app.db.get_all_directions()
        for d in dirs:
            self.tree_dir.insert('', 'end', values=(
                d['id'], d['nom_direction'], d['code_direction'] or '',
                d['responsable'] or '', d['description'] or '', 'Oui' if d['active'] else 'Non'
            ))

    def desactiver_direction(self):
        sel = self.tree_dir.selection()
        if not sel:
            messagebox.showwarning("Attention", "Sélectionnez une direction")
            return
        item = self.tree_dir.item(sel[0])
        dir_id = item['values'][0]
        if messagebox.askyesno("Confirmation", "Désactiver cette direction ?"):
            with self.app.db.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE directions SET active=0 WHERE id=?", (dir_id,))
                conn.commit()
            self.charger_directions()

    # ---------- ACTIVITÉS ----------
    def create_tab_activites(self):
        main = tk.Frame(self.tab_act, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        add_frame = tk.LabelFrame(main, text="Ajouter une activité", font=('Arial',11,'bold'), bg='white')
        add_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(add_frame, text="Nom :", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.act_nom = tk.Entry(add_frame, width=25)
        self.act_nom.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(add_frame, text="Code :", bg='white').grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.act_code = tk.Entry(add_frame, width=15)
        self.act_code.grid(row=0, column=3, padx=5, pady=5)
        tk.Label(add_frame, text="Direction :", bg='white').grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.act_direction = ttk.Combobox(add_frame, width=20)
        self.act_direction.grid(row=0, column=5, padx=5, pady=5)
        tk.Button(add_frame, text="➕ Ajouter", command=self.ajouter_activite,
                  bg='#27ae60', fg='white').grid(row=0, column=6, padx=10, pady=5)

        list_frame = tk.Frame(main, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID', 'Nom', 'Code', 'Direction', 'Responsable', 'Active')
        self.tree_act = ttk.Treeview(list_frame, columns=cols, show='headings', height=12)
        for c in cols:
            self.tree_act.heading(c, text=c)
            self.tree_act.column(c, width=120)
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_act.yview)
        self.tree_act.configure(yscrollcommand=vsb.set)
        self.tree_act.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = tk.Frame(main, bg='white')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btn_frame, text="🗑️ Désactiver", command=self.desactiver_activite,
                  bg='#e74c3c', fg='white').pack(side=tk.RIGHT, padx=5)

        self.charger_directions_combo()

    def charger_directions_combo(self):
        dirs = self.app.db.get_all_directions()
        self.act_direction['values'] = [d['nom_direction'] for d in dirs]

    def ajouter_activite(self):
        nom = self.act_nom.get().strip()
        if not nom:
            messagebox.showwarning("Attention", "Le nom est obligatoire")
            return
        dir_nom = self.act_direction.get()
        dir_id = None
        if dir_nom:
            dirs = self.app.db.get_all_directions()
            for d in dirs:
                if d['nom_direction'] == dir_nom:
                    dir_id = d['id']
                    break
        self.app.db.add_activite(nom, self.act_code.get() or None, dir_id)
        messagebox.showinfo("Succès", f"Activité '{nom}' ajoutée")
        self.act_nom.delete(0, tk.END)
        self.act_code.delete(0, tk.END)
        self.act_direction.set('')
        self.charger_activites()

    def charger_activites(self):
        for i in self.tree_act.get_children():
            self.tree_act.delete(i)
        acts = self.app.db.get_all_activites()
        for a in acts:
            dir_name = ""
            if a['direction_id']:
                dirs = self.app.db.get_all_directions()
                for d in dirs:
                    if d['id'] == a['direction_id']:
                        dir_name = d['nom_direction']
                        break
            self.tree_act.insert('', 'end', values=(
                a['id'], a['nom_activite'], a['code_activite'] or '',
                dir_name, a['responsable'] or '', 'Oui' if a['active'] else 'Non'
            ))

    def desactiver_activite(self):
        sel = self.tree_act.selection()
        if not sel:
            return
        act_id = self.tree_act.item(sel[0])['values'][0]
        with self.app.db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE activites SET active=0 WHERE id=?", (act_id,))
            conn.commit()
        self.charger_activites()

    # ---------- SERVICES ----------
    def create_tab_services(self):
        main = tk.Frame(self.tab_serv, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        add_frame = tk.LabelFrame(main, text="Ajouter un service", font=('Arial',11,'bold'), bg='white')
        add_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(add_frame, text="Nom :", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.serv_nom = tk.Entry(add_frame, width=25)
        self.serv_nom.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(add_frame, text="Code :", bg='white').grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.serv_code = tk.Entry(add_frame, width=15)
        self.serv_code.grid(row=0, column=3, padx=5, pady=5)
        tk.Label(add_frame, text="Activité :", bg='white').grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.serv_activite = ttk.Combobox(add_frame, width=20)
        self.serv_activite.grid(row=0, column=5, padx=5, pady=5)
        tk.Button(add_frame, text="➕ Ajouter", command=self.ajouter_service,
                  bg='#27ae60', fg='white').grid(row=0, column=6, padx=10, pady=5)

        list_frame = tk.Frame(main, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID', 'Nom', 'Code', 'Activité', 'Responsable', 'Active')
        self.tree_serv = ttk.Treeview(list_frame, columns=cols, show='headings', height=12)
        for c in cols:
            self.tree_serv.heading(c, text=c)
            self.tree_serv.column(c, width=120)
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_serv.yview)
        self.tree_serv.configure(yscrollcommand=vsb.set)
        self.tree_serv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = tk.Frame(main, bg='white')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btn_frame, text="🗑️ Désactiver", command=self.desactiver_service,
                  bg='#e74c3c', fg='white').pack(side=tk.RIGHT, padx=5)

        self.charger_activites_combo()

    def charger_activites_combo(self):
        acts = self.app.db.get_all_activites()
        self.serv_activite['values'] = [a['nom_activite'] for a in acts]

    def ajouter_service(self):
        nom = self.serv_nom.get().strip()
        if not nom:
            messagebox.showwarning("Attention", "Le nom est obligatoire")
            return
        act_nom = self.serv_activite.get()
        act_id = None
        if act_nom:
            acts = self.app.db.get_all_activites()
            for a in acts:
                if a['nom_activite'] == act_nom:
                    act_id = a['id']
                    break
        self.app.db.add_service(nom, self.serv_code.get() or None, act_id)
        messagebox.showinfo("Succès", f"Service '{nom}' ajouté")
        self.serv_nom.delete(0, tk.END)
        self.serv_code.delete(0, tk.END)
        self.serv_activite.set('')
        self.charger_services()

    def charger_services(self):
        for i in self.tree_serv.get_children():
            self.tree_serv.delete(i)
        servs = self.app.db.get_all_services()
        for s in servs:
            act_name = ""
            if s['activite_id']:
                acts = self.app.db.get_all_activites()
                for a in acts:
                    if a['id'] == s['activite_id']:
                        act_name = a['nom_activite']
                        break
            self.tree_serv.insert('', 'end', values=(
                s['id'], s['nom_service'], s['code_service'] or '',
                act_name, s['responsable'] or '', 'Oui' if s['active'] else 'Non'
            ))

    def desactiver_service(self):
        sel = self.tree_serv.selection()
        if not sel:
            return
        serv_id = self.tree_serv.item(sel[0])['values'][0]
        with self.app.db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE services SET active=0 WHERE id=?", (serv_id,))
            conn.commit()
        self.charger_services()

    # ---------- ÉQUIPES ----------
    def create_tab_equipes(self):
        main = tk.Frame(self.tab_eq, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        add_frame = tk.LabelFrame(main, text="Ajouter une équipe", font=('Arial',11,'bold'), bg='white')
        add_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(add_frame, text="Nom :", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.eq_nom = tk.Entry(add_frame, width=25)
        self.eq_nom.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(add_frame, text="Code :", bg='white').grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.eq_code = tk.Entry(add_frame, width=15)
        self.eq_code.grid(row=0, column=3, padx=5, pady=5)
        tk.Label(add_frame, text="Service :", bg='white').grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.eq_service = ttk.Combobox(add_frame, width=20)
        self.eq_service.grid(row=0, column=5, padx=5, pady=5)
        tk.Button(add_frame, text="➕ Ajouter", command=self.ajouter_equipe,
                  bg='#27ae60', fg='white').grid(row=0, column=6, padx=10, pady=5)

        list_frame = tk.Frame(main, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('ID', 'Nom', 'Code', 'Service', 'Responsable', 'Active')
        self.tree_eq = ttk.Treeview(list_frame, columns=cols, show='headings', height=12)
        for c in cols:
            self.tree_eq.heading(c, text=c)
            self.tree_eq.column(c, width=120)
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_eq.yview)
        self.tree_eq.configure(yscrollcommand=vsb.set)
        self.tree_eq.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = tk.Frame(main, bg='white')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btn_frame, text="🗑️ Désactiver", command=self.desactiver_equipe,
                  bg='#e74c3c', fg='white').pack(side=tk.RIGHT, padx=5)

        self.charger_services_combo()

    def charger_services_combo(self):
        servs = self.app.db.get_all_services()
        self.eq_service['values'] = [s['nom_service'] for s in servs]

    def ajouter_equipe(self):
        nom = self.eq_nom.get().strip()
        if not nom:
            messagebox.showwarning("Attention", "Le nom est obligatoire")
            return
        serv_nom = self.eq_service.get()
        serv_id = None
        if serv_nom:
            servs = self.app.db.get_all_services()
            for s in servs:
                if s['nom_service'] == serv_nom:
                    serv_id = s['id']
                    break
        self.app.db.add_equipe(nom, self.eq_code.get() or None, serv_id)
        messagebox.showinfo("Succès", f"Équipe '{nom}' ajoutée")
        self.eq_nom.delete(0, tk.END)
        self.eq_code.delete(0, tk.END)
        self.eq_service.set('')
        self.charger_equipes()

    def charger_equipes(self):
        for i in self.tree_eq.get_children():
            self.tree_eq.delete(i)
        eqs = self.app.db.get_all_equipes()
        for e in eqs:
            serv_name = ""
            if e['service_id']:
                servs = self.app.db.get_all_services()
                for s in servs:
                    if s['id'] == e['service_id']:
                        serv_name = s['nom_service']
                        break
            self.tree_eq.insert('', 'end', values=(
                e['id'], e['nom_equipe'], e['code_equipe'] or '',
                serv_name, e['responsable'] or '', 'Oui' if e['active'] else 'Non'
            ))

    def desactiver_equipe(self):
        sel = self.tree_eq.selection()
        if not sel:
            return
        eq_id = self.tree_eq.item(sel[0])['values'][0]
        with self.app.db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE equipes SET active=0 WHERE id=?", (eq_id,))
            conn.commit()
        self.charger_equipes()


# ------------------------------------------------------------
#   IMPORT PERSONNEL
# ------------------------------------------------------------
class ImportPersonnel:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.window = None

    def ouvrir(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Importer du personnel")
        self.window.geometry("700x600")
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.configure(bg='#f0f0f0')
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.window, text="📂 IMPORTATION DE FICHIER", 
                font=('Arial',16,'bold'), bg='#f0f0f0', fg='#2c3e50').pack(pady=10)
        main = tk.Frame(self.window, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        instr = tk.Frame(main, bg='#e8f4f8')
        instr.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(instr, text="📋 Format accepté : CSV ou Excel", 
                font=('Arial',10,'bold'), bg='#e8f4f8').pack(anchor='w', padx=10, pady=5)
        tk.Label(instr, text="Colonnes obligatoires : matricule, nom, prenom, direction", 
                bg='#e8f4f8').pack(anchor='w', padx=20, pady=2)

        sel = tk.Frame(main, bg='white')
        sel.pack(fill=tk.X, padx=10, pady=20)
        tk.Label(sel, text="Fichier source :", font=('Arial',10,'bold'), bg='white').pack(anchor='w', padx=5)
        pth = tk.Frame(sel, bg='white')
        pth.pack(fill=tk.X)
        self.file_path = tk.StringVar()
        tk.Entry(pth, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(pth, text="📂 Parcourir", command=self.parcourir,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=5)

        opt = tk.LabelFrame(main, text="Options", font=('Arial',10,'bold'), bg='white')
        opt.pack(fill=tk.X, padx=10, pady=10)
        self.opt_update = tk.BooleanVar(value=True)
        tk.Checkbutton(opt, text="Mettre à jour les enregistrements existants", 
                      variable=self.opt_update, bg='white').pack(anchor='w', padx=10, pady=5)

        pre = tk.LabelFrame(main, text="Aperçu", font=('Arial',10,'bold'), bg='white')
        pre.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.preview = tk.Text(pre, height=10, font=('Courier',9))
        self.preview.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        btnf = tk.Frame(self.window, bg='#f0f0f0')
        btnf.pack(pady=20)
        tk.Button(btnf, text="✅ IMPORTER", command=self.importer,
                  bg='#27ae60', fg='white', font=('Arial',12,'bold'), width=20, height=2).pack(side=tk.LEFT, padx=10)
        tk.Button(btnf, text="❌ ANNULER", command=self.fermer,
                  bg='#e74c3c', fg='white').pack(side=tk.LEFT, padx=10)

    def parcourir(self):
        fn = filedialog.askopenfilename(
            title="Sélectionner le fichier",
            filetypes=[("CSV","*.csv"), ("Excel","*.xlsx *.xls"), ("Tous","*.*")]
        )
        if fn:
            self.file_path.set(fn)
            self.afficher_apercu(fn)

    def afficher_apercu(self, fn):
        try:
            self.preview.delete(1.0, tk.END)
            if fn.lower().endswith('.csv'):
                df = pd.read_csv(fn, nrows=5)
            else:
                df = pd.read_excel(fn, nrows=5)
            self.preview.insert(tk.END, df.to_string())
            self.preview.insert(tk.END, f"\n\n✅ Fichier: {os.path.basename(fn)}")
            self.preview.insert(tk.END, f"\n📊 Colonnes: {', '.join(df.columns)}")
        except Exception as e:
            self.preview.insert(tk.END, f"❌ Erreur: {e}")

    def importer(self):
        f = self.file_path.get().strip()
        if not f:
            messagebox.showwarning("Attention", "Sélectionnez un fichier")
            return
        if not messagebox.askyesno("Confirmation", "Lancer l'importation ?"):
            return
        self.preview.delete(1.0, tk.END)
        self.preview.insert(tk.END, "⏳ Importation en cours...")
        self.window.update()
        res = self.app.db.import_personnel_from_file(f)
        self.preview.delete(1.0, tk.END)
        msg = f"""
📊 RÉSULTATS
Total: {res['total']}
✅ Importés: {res['importes']}
🔄 Mis à jour: {res['mis_a_jour']}
❌ Erreurs: {res['erreurs']}
"""
        if res['details']:
            msg += "\nDétails:\n" + "\n".join(res['details'][:10])
        self.preview.insert(tk.END, msg)
        if res['erreurs'] == 0:
            messagebox.showinfo("Succès", "Import terminé sans erreur")
            self.fermer()

    def fermer(self):
        if self.window:
            self.window.destroy()


# ------------------------------------------------------------
#   GESTION TOLÉRANCES
# ------------------------------------------------------------
class GestionTolerances:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.window = None
        self.current_agent_id = None

    def ouvrir_interface(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("Gestion des tolérances")
        self.window.geometry("1200x700")
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.configure(bg='#f0f0f0')

        tk.Label(self.window, text="🎯 GESTION DES TOLÉRANCES ET HORAIRES",
                font=('Arial',16,'bold'), bg='#f0f0f0', fg='#2c3e50').pack(pady=10)

        notebook = ttk.Notebook(self.window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_global = ttk.Frame(notebook)
        notebook.add(self.tab_global, text="⚙️ Paramètres globaux")
        self.create_tab_global()

        self.tab_individuel = ttk.Frame(notebook)
        notebook.add(self.tab_individuel, text="👤 Tolérances individuelles")
        self.create_tab_individuel()

        self.tab_analyse = ttk.Frame(notebook)
        notebook.add(self.tab_analyse, text="📊 Analyse des retards")
        self.create_tab_analyse()

    # ---------- Onglet global ----------
    def create_tab_global(self):
        main = tk.Frame(self.tab_global, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        frame_horaires = tk.LabelFrame(main, text="Horaires de travail standard",
                                       font=('Arial',11,'bold'), bg='white')
        frame_horaires.pack(fill=tk.X, padx=10, pady=10)
        grid_h = tk.Frame(frame_horaires, bg='white')
        grid_h.pack(padx=10, pady=10)
        tk.Label(grid_h, text="Heure de début :", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.global_debut = tk.Entry(grid_h, width=10)
        self.global_debut.grid(row=0, column=1, padx=5, pady=5)
        self.global_debut.insert(0, self.app.db.get_parametre('heure_debut_journee', '08:00:00')[:5])
        tk.Label(grid_h, text="Heure de fin :", bg='white').grid(row=0, column=2, padx=15, pady=5, sticky='w')
        self.global_fin = tk.Entry(grid_h, width=10)
        self.global_fin.grid(row=0, column=3, padx=5, pady=5)
        self.global_fin.insert(0, self.app.db.get_parametre('heure_fin_journee', '17:00:00')[:5])
        tk.Label(grid_h, text="Durée de pause :", bg='white').grid(row=0, column=4, padx=15, pady=5, sticky='w')
        self.global_pause = tk.Entry(grid_h, width=10)
        self.global_pause.grid(row=0, column=5, padx=5, pady=5)
        self.global_pause.insert(0, self.app.db.get_parametre('duree_pause', '01:00:00')[:5])

        frame_tol = tk.LabelFrame(main, text="Tolérances par défaut (minutes)",
                                  font=('Arial',11,'bold'), bg='white')
        frame_tol.pack(fill=tk.X, padx=10, pady=10)
        grid_t = tk.Frame(frame_tol, bg='white')
        grid_t.pack(padx=10, pady=10)
        tk.Label(grid_t, text="Retard à l'entrée :", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.global_tol_entree = tk.Spinbox(grid_t, from_=0, to=120, width=8)
        self.global_tol_entree.grid(row=0, column=1, padx=5, pady=5)
        self.global_tol_entree.delete(0, tk.END)
        self.global_tol_entree.insert(0, self.app.db.get_parametre('tolerance_retard_globale', '10'))
        tk.Label(grid_t, text="Départ anticipé :", bg='white').grid(row=0, column=2, padx=15, pady=5, sticky='w')
        self.global_tol_sortie = tk.Spinbox(grid_t, from_=0, to=120, width=8)
        self.global_tol_sortie.grid(row=0, column=3, padx=5, pady=5)
        self.global_tol_sortie.delete(0, tk.END)
        self.global_tol_sortie.insert(0, self.app.db.get_parametre('tolerance_depart_anticipe', '10'))

        frame_policy = tk.LabelFrame(main, text="Politique de gestion des retards",
                                     font=('Arial',11,'bold'), bg='white')
        frame_policy.pack(fill=tk.X, padx=10, pady=10)
        grid_p = tk.Frame(frame_policy, bg='white')
        grid_p.pack(padx=10, pady=10)
        tk.Label(grid_p, text="Seuil de justification (minutes) :", bg='white').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.seuil_justif = tk.Spinbox(grid_p, from_=0, to=240, width=8)
        self.seuil_justif.grid(row=0, column=1, padx=5, pady=5)
        self.seuil_justif.delete(0, tk.END)
        self.seuil_justif.insert(0, self.app.db.get_parametre('seuil_justification_retard', '15'))
        tk.Label(grid_p, text="Pénalité par minute non justifiée (€) :", bg='white').grid(row=0, column=2, padx=15, pady=5, sticky='w')
        self.penalite = tk.Entry(grid_p, width=8)
        self.penalite.grid(row=0, column=3, padx=5, pady=5)
        self.penalite.insert(0, self.app.db.get_parametre('penalite_retard', '0.50'))
        self.arrondi_var = tk.BooleanVar(value=self.app.db.get_parametre('arrondir_retard', '0') == '1')
        tk.Checkbutton(grid_p, text="Arrondir les retards à 5 minutes supérieures",
                       variable=self.arrondi_var, bg='white').grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky='w')

        btn_frame = tk.Frame(main, bg='white')
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="💾 Sauvegarder les paramètres globaux",
                  command=self.sauvegarder_parametres_globaux,
                  bg='#27ae60', fg='white', font=('Arial',11,'bold'),
                  width=30, height=2).pack()
        tk.Button(btn_frame, text="🔄 Appliquer à tous les agents",
                  command=self.appliquer_a_tous,
                  bg='#3498db', fg='white', font=('Arial',10),
                  width=25, height=1).pack(pady=5)

    def sauvegarder_parametres_globaux(self):
        try:
            self.app.db.set_parametre('heure_debut_journee', self.global_debut.get().strip() + ':00')
            self.app.db.set_parametre('heure_fin_journee', self.global_fin.get().strip() + ':00')
            self.app.db.set_parametre('duree_pause', self.global_pause.get().strip() + ':00')
            self.app.db.set_parametre('tolerance_retard_globale', self.global_tol_entree.get().strip())
            self.app.db.set_parametre('tolerance_depart_anticipe', self.global_tol_sortie.get().strip())
            self.app.db.set_parametre('seuil_justification_retard', self.seuil_justif.get().strip())
            self.app.db.set_parametre('penalite_retard', self.penalite.get().strip())
            self.app.db.set_parametre('arrondir_retard', '1' if self.arrondi_var.get() else '0')
            messagebox.showinfo("Succès", "Paramètres globaux sauvegardés")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde : {e}")

    def appliquer_a_tous(self):
        if messagebox.askyesno("Confirmation", "Voulez-vous appliquer les tolérances par défaut à TOUS les agents ?\nCette action écrasera leurs tolérances individuelles actuelles."):
            try:
                tol_entree = int(self.global_tol_entree.get())
                tol_sortie = int(self.global_tol_sortie.get())
                heure_entree = self.global_debut.get().strip() + ':00'
                heure_sortie = self.global_fin.get().strip() + ':00'
                with self.app.db.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute('''
                        UPDATE personnel SET
                            heure_entree_theorique = ?,
                            heure_sortie_theorique = ?,
                            tolerance_entree = ?,
                            tolerance_sortie = ?,
                            updated_at = ?
                    ''', (heure_entree, heure_sortie, tol_entree, tol_sortie, datetime.now()))
                    conn.commit()
                messagebox.showinfo("Succès", f"{cursor.rowcount} agents mis à jour.")
                if hasattr(self, 'tree_agents'):
                    self.charger_liste_agents()
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'application : {e}")

    # ---------- Onglet individuel ----------
    def create_tab_individuel(self):
        main = tk.Frame(self.tab_individuel, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        search_frame = tk.Frame(main, bg='white')
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(search_frame, text="Rechercher un agent :", bg='white', font=('Arial',10,'bold')).pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind('<Return>', lambda e: self.charger_liste_agents())
        tk.Button(search_frame, text="🔍", command=self.charger_liste_agents,
                  bg='#3498db', fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(search_frame, text="🔄 Réinitialiser", command=self.reset_recherche,
                  bg='#95a5a6', fg='white').pack(side=tk.LEFT, padx=5)

        content_frame = tk.Frame(main, bg='white')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        list_frame = tk.Frame(content_frame, bg='white')
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        cols = ('ID', 'Matricule', 'Nom', 'Prénom', 'Entrée th.', 'Sortie th.', 'Tol. E', 'Tol. S', 'Quart')
        self.tree_agents = ttk.Treeview(list_frame, columns=cols, show='headings', height=20)
        col_widths = [50,100,120,120,100,100,80,80,100]
        for c,w in zip(cols,col_widths):
            self.tree_agents.heading(c, text=c)
            self.tree_agents.column(c, width=w)
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_agents.yview)
        self.tree_agents.configure(yscrollcommand=vsb.set)
        self.tree_agents.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_agents.bind('<<TreeviewSelect>>', self.on_agent_select)

        edit_frame = tk.LabelFrame(content_frame, text="Modifier les tolérances",
                                   font=('Arial',11,'bold'), bg='white')
        edit_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10,0))
        info_frame = tk.Frame(edit_frame, bg='#ecf0f1', relief=tk.SUNKEN, bd=1)
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        self.agent_info = tk.Label(info_frame, text="Aucun agent sélectionné",
                                   font=('Arial',10,'bold'), bg='#ecf0f1', anchor='w')
        self.agent_info.pack(padx=5, pady=5, fill=tk.X)

        form_f = tk.Frame(edit_frame, bg='white')
        form_f.pack(padx=10, pady=10, fill=tk.X)
        row = 0
        tk.Label(form_f, text="Heure d'entrée théorique :", bg='white').grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.entree_h = tk.Spinbox(form_f, from_=0, to=23, width=3, format='%02.0f')
        self.entree_h.grid(row=row, column=1, padx=2, pady=5)
        tk.Label(form_f, text=":", bg='white').grid(row=row, column=2)
        self.entree_m = tk.Spinbox(form_f, from_=0, to=59, width=3, format='%02.0f')
        self.entree_m.grid(row=row, column=3, padx=2, pady=5)
        tk.Label(form_f, text=":00", bg='white').grid(row=row, column=4)
        row += 1
        tk.Label(form_f, text="Heure de sortie théorique :", bg='white').grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.sortie_h = tk.Spinbox(form_f, from_=0, to=23, width=3, format='%02.0f')
        self.sortie_h.grid(row=row, column=1, padx=2, pady=5)
        tk.Label(form_f, text=":", bg='white').grid(row=row, column=2)
        self.sortie_m = tk.Spinbox(form_f, from_=0, to=59, width=3, format='%02.0f')
        self.sortie_m.grid(row=row, column=3, padx=2, pady=5)
        tk.Label(form_f, text=":00", bg='white').grid(row=row, column=4)
        row += 1
        tk.Label(form_f, text="Tolérance entrée (min) :", bg='white').grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.tol_entree_var = tk.IntVar()
        self.tol_entree_spin = tk.Spinbox(form_f, from_=0, to=120, textvariable=self.tol_entree_var, width=8)
        self.tol_entree_spin.grid(row=row, column=1, padx=5, pady=5, sticky='w')
        row += 1
        tk.Label(form_f, text="Tolérance sortie (min) :", bg='white').grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.tol_sortie_var = tk.IntVar()
        self.tol_sortie_spin = tk.Spinbox(form_f, from_=0, to=120, textvariable=self.tol_sortie_var, width=8)
        self.tol_sortie_spin.grid(row=row, column=1, padx=5, pady=5, sticky='w')
        row += 1
        tk.Label(form_f, text="Type de quart :", bg='white').grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.quart_var = tk.StringVar()
        self.quart_combo = ttk.Combobox(form_f, textvariable=self.quart_var,
                                         values=['jour','nuit','personnalisé'], width=15)
        self.quart_combo.grid(row=row, column=1, padx=5, pady=5, sticky='w')
        row += 1
        self.concerne_var = tk.BooleanVar()
        tk.Checkbutton(form_f, text="Agent concerné par le pointage",
                       variable=self.concerne_var, bg='white').grid(row=row, column=0, columnspan=2, padx=5, pady=5, sticky='w')

        btn_frame = tk.Frame(edit_frame, bg='white')
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="💾 Mettre à jour", command=self.mettre_a_jour_agent,
                  bg='#27ae60', fg='white', font=('Arial',10,'bold'),
                  width=20, height=1).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="↩ Réinitialiser", command=self.reset_formulaire,
                  bg='#f39c12', fg='white', width=15).pack(side=tk.LEFT, padx=5)

        self.charger_liste_agents()

    def charger_liste_agents(self):
        for i in self.tree_agents.get_children():
            self.tree_agents.delete(i)
        search = self.search_var.get().strip()
        if search:
            agents = self.app.db.search_personnel(search)
        else:
            agents = self.app.db.get_personnel()
        for a in agents:
            entree = a['heure_entree_theorique'][:5] if a['heure_entree_theorique'] else '08:00'
            sortie = a['heure_sortie_theorique'][:5] if a['heure_sortie_theorique'] else '17:00'
            self.tree_agents.insert('','end', values=(
                a['id'],
                a['matricule'],
                a['nom'],
                a['prenom'],
                entree,
                sortie,
                a['tolerance_entree'],
                a['tolerance_sortie'],
                a['type_quart'] or 'jour'
            ))

    def reset_recherche(self):
        self.search_var.set('')
        self.charger_liste_agents()

    def on_agent_select(self, event):
        sel = self.tree_agents.selection()
        if not sel:
            return
        item = self.tree_agents.item(sel[0])
        vals = item['values']
        self.current_agent_id = vals[0]
        mat = vals[1]
        nom = vals[2]
        prenom = vals[3]
        agent = self.app.db.get_personnel(self.current_agent_id)
        if not agent:
            return
        self.agent_info.config(text=f"Agent sélectionné : {prenom} {nom} (Matricule: {mat})")
        entree = agent['heure_entree_theorique'] or '08:00:00'
        sortie = agent['heure_sortie_theorique'] or '17:00:00'
        self.entree_h.delete(0,tk.END); self.entree_h.insert(0, entree[:2])
        self.entree_m.delete(0,tk.END); self.entree_m.insert(0, entree[3:5])
        self.sortie_h.delete(0,tk.END); self.sortie_h.insert(0, sortie[:2])
        self.sortie_m.delete(0,tk.END); self.sortie_m.insert(0, sortie[3:5])
        self.tol_entree_var.set(agent['tolerance_entree'] or 10)
        self.tol_sortie_var.set(agent['tolerance_sortie'] or 10)
        self.quart_var.set(agent['type_quart'] or 'jour')
        self.concerne_var.set(bool(agent['concerne_pointage']))

    def mettre_a_jour_agent(self):
        if not self.current_agent_id:
            messagebox.showwarning("Attention", "Veuillez sélectionner un agent")
            return
        try:
            heure_entree = f"{int(self.entree_h.get()):02d}:{int(self.entree_m.get()):02d}:00"
            heure_sortie = f"{int(self.sortie_h.get()):02d}:{int(self.sortie_m.get()):02d}:00"
            tol_entree = self.tol_entree_var.get()
            tol_sortie = self.tol_sortie_var.get()
            type_quart = self.quart_var.get()
            concerne = 1 if self.concerne_var.get() else 0
            agent = self.app.db.get_personnel(self.current_agent_id)
            if not agent:
                return
            data = (
                agent['matricule'],
                agent['badge_id'],
                agent['nom'],
                agent['prenom'],
                agent['type_person'],
                agent['fonction'],
                agent['activite'],
                agent['division'],
                agent['direction'],
                agent['departement'],
                agent['service'],
                agent['equipe'],
                agent['date_embauche'],
                agent['date_naissance'],
                agent['adresse'],
                agent['telephone'],
                agent['email'],
                agent['photo'],
                agent['statut'],
                concerne,
                type_quart,
                heure_entree,
                heure_sortie,
                tol_entree,
                tol_sortie
            )
            self.app.db.update_personnel(self.current_agent_id, data)
            messagebox.showinfo("Succès", "Tolérances mises à jour")
            self.charger_liste_agents()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la mise à jour : {e}")

    def reset_formulaire(self):
        if self.current_agent_id:
            agent = self.app.db.get_personnel(self.current_agent_id)
            if agent:
                entree = agent['heure_entree_theorique'] or '08:00:00'
                sortie = agent['heure_sortie_theorique'] or '17:00:00'
                self.entree_h.delete(0,tk.END); self.entree_h.insert(0, entree[:2])
                self.entree_m.delete(0,tk.END); self.entree_m.insert(0, entree[3:5])
                self.sortie_h.delete(0,tk.END); self.sortie_h.insert(0, sortie[:2])
                self.sortie_m.delete(0,tk.END); self.sortie_m.insert(0, sortie[3:5])
                self.tol_entree_var.set(agent['tolerance_entree'] or 10)
                self.tol_sortie_var.set(agent['tolerance_sortie'] or 10)
                self.quart_var.set(agent['type_quart'] or 'jour')
                self.concerne_var.set(bool(agent['concerne_pointage']))

    # ---------- Onglet analyse ----------
    def create_tab_analyse(self):
        main = tk.Frame(self.tab_analyse, bg='white', relief=tk.RAISED, bd=2)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        filter_frame = tk.Frame(main, bg='white')
        filter_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(filter_frame, text="Période du :", bg='white').grid(row=0, column=0, padx=5, pady=5)
        self.analyse_date_debut = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd')
        self.analyse_date_debut.grid(row=0, column=1, padx=5, pady=5)
        self.analyse_date_debut.set_date(date.today().replace(day=1))
        tk.Label(filter_frame, text="au :", bg='white').grid(row=0, column=2, padx=5, pady=5)
        self.analyse_date_fin = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd')
        self.analyse_date_fin.grid(row=0, column=3, padx=5, pady=5)
        self.analyse_date_fin.set_date(date.today())
        tk.Label(filter_frame, text="Direction :", bg='white').grid(row=0, column=4, padx=5, pady=5)
        self.analyse_direction = ttk.Combobox(filter_frame, width=20)
        self.analyse_direction.grid(row=0, column=5, padx=5, pady=5)
        self.charger_directions_analyse()
        tk.Button(filter_frame, text="📊 Analyser", command=self.analyser_retards,
                  bg='#3498db', fg='white').grid(row=0, column=6, padx=10, pady=5)

        result_frame = tk.Frame(main, bg='white')
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ('Matricule', 'Agent', 'Nb retards', 'Total minutes', 'Justifiés', 'Non justifiés', 'Pénalité (€)')
        self.tree_analyse = ttk.Treeview(result_frame, columns=cols, show='headings', height=12)
        for c in cols:
            self.tree_analyse.heading(c, text=c)
            self.tree_analyse.column(c, width=120)
        vsb = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree_analyse.yview)
        self.tree_analyse.configure(yscrollcommand=vsb.set)
        self.tree_analyse.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        btn_export = tk.Frame(main, bg='white')
        btn_export.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btn_export, text="📤 Exporter CSV", command=self.exporter_analyse,
                  bg='#2ecc71', fg='white').pack(side=tk.RIGHT, padx=5)

    def charger_directions_analyse(self):
        try:
            dirs = self.app.db.get_all_directions()
            self.analyse_direction['values'] = ['Toutes'] + [d['nom_direction'] for d in dirs]
            self.analyse_direction.set('Toutes')
        except:
            pass

    def analyser_retards(self):
        for i in self.tree_analyse.get_children():
            self.tree_analyse.delete(i)
        date_debut = self.analyse_date_debut.get_date().strftime('%Y-%m-%d')
        date_fin = self.analyse_date_fin.get_date().strftime('%Y-%m-%d')
        direction = None if self.analyse_direction.get() == 'Toutes' else self.analyse_direction.get()
        try:
            with self.app.db.get_connection() as conn:
                cursor = conn.cursor()
                query = '''
                    SELECT r.matricule, p.nom, p.prenom,
                        COUNT(r.id) as nb_retards,
                        SUM(r.minutes_retard) as total_min,
                        SUM(CASE WHEN r.est_justifie=1 THEN r.minutes_retard ELSE 0 END) as justifie,
                        SUM(CASE WHEN r.est_justifie=0 THEN r.minutes_retard ELSE 0 END) as non_justifie
                    FROM retards_cumules r
                    JOIN personnel p ON r.personnel_id = p.id
                    WHERE r.date_retard BETWEEN ? AND ?
                '''
                params = [date_debut, date_fin]
                if direction:
                    query += ' AND p.direction = ?'
                    params.append(direction)
                query += ' GROUP BY r.personnel_id ORDER BY total_min DESC'
                cursor.execute(query, params)
                rows = cursor.fetchall()
                penalite_par_min = float(self.app.db.get_parametre('penalite_retard', '0.50'))
                for r in rows:
                    penalite = r['non_justifie'] * penalite_par_min
                    self.tree_analyse.insert('', 'end', values=(
                        r['matricule'],
                        f"{r['prenom']} {r['nom']}",
                        r['nb_retards'],
                        f"{r['total_min']} min",
                        f"{r['justifie']} min",
                        f"{r['non_justifie']} min",
                        f"{penalite:.2f}"
                    ))
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'analyse : {e}")

    def exporter_analyse(self):
        data = []
        for i in self.tree_analyse.get_children():
            data.append(self.tree_analyse.item(i)['values'])
        if not data:
            messagebox.showwarning("Attention", "Aucune donnée à exporter")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".csv",
                                          filetypes=[("CSV","*.csv"), ("Excel","*.xlsx")])
        if fn:
            try:
                df = pd.DataFrame(data, columns=['Matricule','Agent','Nb retards','Total minutes','Justifiés','Non justifiés','Pénalité (€)'])
                if fn.endswith('.csv'):
                    df.to_csv(fn, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(fn, index=False)
                messagebox.showinfo("Succès", "Export terminé")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur export: {e}")


# ------------------------------------------------------------
#   PERSONNEL DIALOG (Ajout / Modification)
# ------------------------------------------------------------
class PersonnelDialog:
    def __init__(self, parent, app, personnel_id=None):
        self.app = app
        self.personnel_id = personnel_id
        self.result = False
        
        # ...
        self.personnel = None
        if personnel_id:
           self.personnel = app.db.get_personnel(personnel_id)
        # ...

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Modifier le personnel" if personnel_id else "Ajouter du personnel")
        self.dialog.geometry("800x850")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.canvas = tk.Canvas(self.dialog, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self.dialog, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='white')

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.create_widgets()
        if personnel_id:
            self.load_data()
        else:
            self.concerne_var.set(True)
            self.statut_var.set('actif')
            self.type_quart_var.set('jour')
            self.heure_entree_var.set('08:00:00')
            self.heure_sortie_var.set('17:00:00')
            self.tolerance_entree_var.set(10)
            self.tolerance_sortie_var.set(10)

        self.dialog.wait_window()

    def create_widgets(self):
        form = self.scrollable_frame
        row = 0

        # --------------------------------------------------------
        # SECTION IDENTIFICATION
        # --------------------------------------------------------
        tk.Label(form, text="IDENTIFICATION", font=('Arial',12,'bold'),
                 fg='#3498db', bg='white').grid(row=row, column=0, columnspan=4,
                                                sticky='w', pady=(10,5), padx=10)
        row += 1

        tk.Label(form, text="Matricule *:", bg='white').grid(row=row, column=0,
                                                              sticky='w', pady=5, padx=20)
        self.matricule_var = tk.StringVar()
        self.entry_matricule = tk.Entry(form, textvariable=self.matricule_var, width=20)
        self.entry_matricule.grid(row=row, column=1, pady=5, padx=10)
        if self.personnel_id:
            self.entry_matricule.config(state='readonly')
        row += 1

        tk.Label(form, text="Badge ID :", bg='white').grid(row=row, column=0,
                                                           sticky='w', pady=5, padx=20)
        self.badge_var = tk.StringVar()
        tk.Entry(form, textvariable=self.badge_var, width=20).grid(row=row, column=1,
                                                                    pady=5, padx=10)
        row += 1

        tk.Label(form, text="Nom *:", bg='white').grid(row=row, column=0,
                                                       sticky='w', pady=5, padx=20)
        self.nom_var = tk.StringVar()
        tk.Entry(form, textvariable=self.nom_var, width=20).grid(row=row, column=1,
                                                                  pady=5, padx=10)
        row += 1

        tk.Label(form, text="Prénom *:", bg='white').grid(row=row, column=0,
                                                          sticky='w', pady=5, padx=20)
        self.prenom_var = tk.StringVar()
        tk.Entry(form, textvariable=self.prenom_var, width=20).grid(row=row, column=1,
                                                                    pady=5, padx=10)
        row += 1

        tk.Label(form, text="Type de personnel *:", bg='white').grid(row=row, column=0,
                                                                     sticky='w', pady=5, padx=20)
        self.type_person_var = tk.StringVar(value='cadre')
        types = ['cadre_dirigeant', 'cadre_superieur', 'cadre', 'maitrise',
                 'technicien', 'employe', 'ouvrier', 'stagiaire']
        self.type_combo = ttk.Combobox(form, textvariable=self.type_person_var,
                                       values=types, width=18, state='readonly')
        self.type_combo.grid(row=row, column=1, pady=5, padx=10)
        row += 1

        tk.Label(form, text="Date de naissance :", bg='white').grid(row=row, column=0,
                                                                    sticky='w', pady=5, padx=20)
        self.date_naissance = DateEntry(form, width=17, date_pattern='yyyy-mm-dd',
                                        background='white', foreground='black')
        self.date_naissance.grid(row=row, column=1, pady=5, padx=10)
        row += 1
        


        # --------------------------------------------------------
        # SECTION FONCTION & HIÉRARCHIE
        # --------------------------------------------------------
        tk.Label(form, text="FONCTION & HIÉRARCHIE", font=('Arial',12,'bold'),
                 fg='#3498db', bg='white').grid(row=row, column=0, columnspan=4,
                                                sticky='w', pady=(15,5), padx=10)
        row += 1

        tk.Label(form, text="Fonction *:", bg='white').grid(row=row, column=0,
                                                            sticky='w', pady=5, padx=20)
        self.fonction_var = tk.StringVar()
        tk.Entry(form, textvariable=self.fonction_var, width=20).grid(row=row, column=1,
                                                                      pady=5, padx=10)
        row += 1

        tk.Label(form, text="Activité :", bg='white').grid(row=row, column=0,
                                                           sticky='w', pady=5, padx=20)
        self.activite_var = tk.StringVar()
        self.activite_combo = ttk.Combobox(form, textvariable=self.activite_var, width=18)
        self.activite_combo.grid(row=row, column=1, pady=5, padx=10)
        row += 1

        tk.Label(form, text="Division :", bg='white').grid(row=row, column=0,
                                                           sticky='w', pady=5, padx=20)
        self.division_var = tk.StringVar()
        tk.Entry(form, textvariable=self.division_var, width=20).grid(row=row, column=1,
                                                                      pady=5, padx=10)
        row += 1

        tk.Label(form, text="Direction *:", bg='white').grid(row=row, column=0,
                                                             sticky='w', pady=5, padx=20)
        self.direction_var = tk.StringVar()
        self.direction_combo = ttk.Combobox(form, textvariable=self.direction_var, width=18)
        self.direction_combo.grid(row=row, column=1, pady=5, padx=10)
        row += 1

        tk.Label(form, text="Département :", bg='white').grid(row=row, column=0,
                                                              sticky='w', pady=5, padx=20)
        self.departement_var = tk.StringVar()
        tk.Entry(form, textvariable=self.departement_var, width=20).grid(row=row, column=1,
                                                                         pady=5, padx=10)
        row += 1

        tk.Label(form, text="Service :", bg='white').grid(row=row, column=0,
                                                          sticky='w', pady=5, padx=20)
        self.service_var = tk.StringVar()
        self.service_combo = ttk.Combobox(form, textvariable=self.service_var, width=18)
        self.service_combo.grid(row=row, column=1, pady=5, padx=10)
        row += 1

        tk.Label(form, text="Équipe :", bg='white').grid(row=row, column=0,
                                                         sticky='w', pady=5, padx=20)
        self.equipe_var = tk.StringVar()
        self.equipe_combo = ttk.Combobox(form, textvariable=self.equipe_var, width=18)
        self.equipe_combo.grid(row=row, column=1, pady=5, padx=10)
        row += 1

        # --------------------------------------------------------
        # SECTION CONTRAT & STATUT
        # --------------------------------------------------------
        tk.Label(form, text="CONTRAT & STATUT", font=('Arial',12,'bold'),
                 fg='#3498db', bg='white').grid(row=row, column=0, columnspan=4,
                                                sticky='w', pady=(15,5), padx=10)
        row += 1

        tk.Label(form, text="Date d'embauche :", bg='white').grid(row=row, column=0,
                                                                  sticky='w', pady=5, padx=20)
        self.date_embauche = DateEntry(form, width=17, date_pattern='yyyy-mm-dd',
                                       background='white', foreground='black')
        self.date_embauche.grid(row=row, column=1, pady=5, padx=10)
        row += 1

        tk.Label(form, text="Statut :", bg='white').grid(row=row, column=0,
                                                         sticky='w', pady=5, padx=20)
        self.statut_var = tk.StringVar(value='actif')
        self.statut_combo = ttk.Combobox(form, textvariable=self.statut_var,
                                         values=['actif', 'inactif', 'congé'],
                                         width=17, state='readonly')
        self.statut_combo.grid(row=row, column=1, pady=5, padx=10)
        row += 1

        tk.Label(form, text="Concerné par le pointage :", bg='white').grid(row=row, column=0,
                                                                           sticky='w', pady=5, padx=20)
        self.concerne_var = tk.BooleanVar(value=True)
        tk.Checkbutton(form, variable=self.concerne_var, bg='white').grid(row=row, column=1,
                                                                          sticky='w', pady=5, padx=10)
        row += 1

        tk.Label(form, text="Type de quart :", bg='white').grid(row=row, column=0,
                                                                sticky='w', pady=5, padx=20)
        self.type_quart_var = tk.StringVar(value='jour')
        self.quart_combo = ttk.Combobox(form, textvariable=self.type_quart_var,
                                        values=['jour', 'nuit', 'personnalisé'],
                                        width=17, state='readonly')
        self.quart_combo.grid(row=row, column=1, pady=5, padx=10)
        row += 1

        # --------------------------------------------------------
        # SECTION HORAIRES & TOLÉRANCES
        # --------------------------------------------------------
        tk.Label(form, text="HORAIRES & TOLÉRANCES", font=('Arial',12,'bold'),
                 fg='#3498db', bg='white').grid(row=row, column=0, columnspan=4,
                                                sticky='w', pady=(15,5), padx=10)
        row += 1

        tk.Label(form, text="Heure d'entrée théorique (HH:MM:SS) :", bg='white').grid(
            row=row, column=0, sticky='w', pady=5, padx=20)
        self.heure_entree_var = tk.StringVar(value='08:00:00')
        tk.Entry(form, textvariable=self.heure_entree_var, width=15).grid(
            row=row, column=1, sticky='w', pady=5, padx=10)
        row += 1

        tk.Label(form, text="Heure de sortie théorique (HH:MM:SS) :", bg='white').grid(
            row=row, column=0, sticky='w', pady=5, padx=20)
        self.heure_sortie_var = tk.StringVar(value='17:00:00')
        tk.Entry(form, textvariable=self.heure_sortie_var, width=15).grid(
            row=row, column=1, sticky='w', pady=5, padx=10)
        row += 1

        tk.Label(form, text="Tolérance à l'entrée (minutes) :", bg='white').grid(
            row=row, column=0, sticky='w', pady=5, padx=20)
        self.tolerance_entree_var = tk.IntVar(value=10)
        tk.Spinbox(form, from_=0, to=120, textvariable=self.tolerance_entree_var,
                   width=10).grid(row=row, column=1, sticky='w', pady=5, padx=10)
        row += 1

        tk.Label(form, text="Tolérance à la sortie (minutes) :", bg='white').grid(
            row=row, column=0, sticky='w', pady=5, padx=20)
        self.tolerance_sortie_var = tk.IntVar(value=10)
        tk.Spinbox(form, from_=0, to=120, textvariable=self.tolerance_sortie_var,
                   width=10).grid(row=row, column=1, sticky='w', pady=5, padx=10)
        row += 1

        # --------------------------------------------------------
        # SECTION COORDONNÉES
        # --------------------------------------------------------
        tk.Label(form, text="COORDONNÉES", font=('Arial',12,'bold'),
                 fg='#3498db', bg='white').grid(row=row, column=0, columnspan=4,
                                                sticky='w', pady=(15,5), padx=10)
        row += 1

        tk.Label(form, text="Adresse :", bg='white').grid(row=row, column=0,
                                                          sticky='w', pady=5, padx=20)
        self.adresse_var = tk.StringVar()
        tk.Entry(form, textvariable=self.adresse_var, width=30).grid(
            row=row, column=1, columnspan=2, sticky='w', pady=5, padx=10)
        row += 1

        tk.Label(form, text="Téléphone :", bg='white').grid(row=row, column=0,
                                                            sticky='w', pady=5, padx=20)
        self.telephone_var = tk.StringVar()
        tk.Entry(form, textvariable=self.telephone_var, width=20).grid(
            row=row, column=1, sticky='w', pady=5, padx=10)
        row += 1

        tk.Label(form, text="Email :", bg='white').grid(row=row, column=2,
                                                        sticky='w', pady=5, padx=20)
        self.email_var = tk.StringVar()
        tk.Entry(form, textvariable=self.email_var, width=25).grid(
            row=row, column=3, sticky='w', pady=5, padx=10)
        row += 1

        self.photo_var = tk.StringVar(value=None)
        
        # --------------------------------------------------------
                # SECTION PHOTO
        # --------------------------------------------------------       
                # Après la section COORDONNÉES (ou avant les boutons), ajoutez :
        tk.Label(form, text="PHOTO", font=('Arial', 12, 'bold'),
                 fg='#3498db').grid(row=row, column=0, columnspan=4, sticky='w', pady=(10,5))
        row += 1

        # Cadre pour la photo
        photo_frame = tk.Frame(form, bg='white')
        photo_frame.grid(row=row, column=0, columnspan=4, pady=10)

        self.photo_canvas = tk.Canvas(photo_frame, width=150, height=150, bg='#f0f0f0', relief='sunken', bd=2)
        self.photo_canvas.pack(side=tk.LEFT, padx=10)

        btn_photo = tk.Button(photo_frame, text="📷 Choisir une photo",
                              command=self.choisir_photo,
                              bg='#3498db', fg='white')
        btn_photo.pack(side=tk.LEFT, padx=10)

        row += 1

        # --------------------------------------------------------
        # BOUTONS
        # --------------------------------------------------------
        btn_frame = tk.Frame(form, bg='white')
        btn_frame.grid(row=row, column=0, columnspan=4, pady=30)

        tk.Button(btn_frame, text="💾 ENREGISTRER", command=self.save,
                  bg='#27ae60', fg='white', font=('Arial',11,'bold'),
                  width=20, height=2).pack(side=tk.LEFT, padx=10)

        tk.Button(btn_frame, text="❌ ANNULER", command=self.dialog.destroy,
                  bg='#e74c3c', fg='white', font=('Arial',10),
                  width=15, height=1).pack(side=tk.LEFT, padx=10)

        self.charger_listes()

    def charger_listes(self):
        try:
            directions = self.app.db.get_all_directions()
            self.direction_combo['values'] = [d['nom_direction'] for d in directions]
            activites = self.app.db.get_all_activites()
            self.activite_combo['values'] = [a['nom_activite'] for a in activites]
            services = self.app.db.get_all_services()
            self.service_combo['values'] = [s['nom_service'] for s in services]
            equipes = self.app.db.get_all_equipes()
            self.equipe_combo['values'] = [e['nom_equipe'] for e in equipes]
        except Exception as e:
            print(f"⚠️ Erreur chargement listes: {e}")
            
            
    def choisir_photo(self):
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Sélectionner une photo",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif")]
        )
        if file_path:
            self.photo_path = file_path
            self.afficher_photo(file_path)

    def afficher_photo(self, path):
        
        try:
            img = Image.open(path)
        # Taille du canvas
            canvas_width = 150
            canvas_height = 150
        # Redimensionner en conservant les proportions
            img.thumbnail((canvas_width, canvas_height), Image.LANCZOS)
            self.photo_image = ImageTk.PhotoImage(img)
        # Effacer le contenu précédent
            self.photo_canvas.delete("all")
        # Calculer la position pour centrer
            x = (canvas_width - img.width) // 2
            y = (canvas_height - img.height) // 2
            self.photo_canvas.create_image(x, y, anchor=tk.NW, image=self.photo_image)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger l'image : {e}")
            
    def load_data(self):
        personnel = self.app.db.get_personnel(self.personnel_id)
        if not personnel:
            return
        self.matricule_var.set(personnel['matricule'])
        self.badge_var.set(personnel['badge_id'] or '')
        self.nom_var.set(personnel['nom'])
        self.prenom_var.set(personnel['prenom'])
        self.type_person_var.set(personnel['type_person'] or 'cadre')
        if personnel['date_naissance']:
            self.date_naissance.set_date(datetime.strptime(personnel['date_naissance'], '%Y-%m-%d'))
        self.fonction_var.set(personnel['fonction'])
        self.activite_var.set(personnel['activite'] or '')
        self.division_var.set(personnel['division'] or '')
        self.direction_var.set(personnel['direction'])
        self.departement_var.set(personnel['departement'] or '')
        self.service_var.set(personnel['service'] or '')
        self.equipe_var.set(personnel['equipe'] or '')
        if personnel['date_embauche']:
            self.date_embauche.set_date(datetime.strptime(personnel['date_embauche'], '%Y-%m-%d'))
        self.statut_var.set(personnel['statut'])
        self.concerne_var.set(bool(personnel['concerne_pointage']))
        self.type_quart_var.set(personnel['type_quart'] or 'jour')
        self.heure_entree_var.set(personnel['heure_entree_theorique'] or '08:00:00')
        self.heure_sortie_var.set(personnel['heure_sortie_theorique'] or '17:00:00')
        self.tolerance_entree_var.set(personnel['tolerance_entree'] or 10)
        self.tolerance_sortie_var.set(personnel['tolerance_sortie'] or 10)
        self.adresse_var.set(personnel['adresse'] or '')
        self.telephone_var.set(personnel['telephone'] or '')
        self.email_var.set(personnel['email'] or '')
        self.photo_actuelle = personnel['photo']
        if personnel['photo']:
            self.photo_path = personnel['photo']
            self.afficher_photo(personnel['photo'])           

    def save(self):
        
        # Validation des champs obligatoires
        if not all([self.matricule_var.get(), self.nom_var.get(), self.prenom_var.get(),
               self.fonction_var.get(), self.direction_var.get()]):
            messagebox.showwarning("Validation", "Veuillez remplir tous les champs obligatoires (*)")
            return
    
        photo = None
        # Gestion de la photo
        if hasattr(self, 'photo_path') and self.photo_path:
        # Sauvegarder la nouvelle photo
           dossier = os.path.join('pointage_data', 'photos')
           os.makedirs(dossier, exist_ok=True)
           ext = os.path.splitext(self.photo_path)[1]
           nom_fichier = f"photo_{self.matricule_var.get()}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
           chemin = os.path.join(dossier, nom_fichier)
        # Redimensionner et sauvegarder
           img = Image.open(self.photo_path)
           img.save(chemin)
           photo = chemin
        else:
        # Pas de nouvelle photo : on garde l'ancienne (si modification) ou None (si ajout)
           photo = self.photo_actuelle if hasattr(self, 'photo_actuelle') else None
       
        if not self.matricule_var.get().strip():
            messagebox.showwarning("Validation", "Le matricule est obligatoire.")
            return
        if not self.nom_var.get().strip():
            messagebox.showwarning("Validation", "Le nom est obligatoire.")
            return
        if not self.prenom_var.get().strip():
            messagebox.showwarning("Validation", "Le prénom est obligatoire.")
            return
        if not self.fonction_var.get().strip():
            messagebox.showwarning("Validation", "La fonction est obligatoire.")
            return
        if not self.direction_var.get().strip():
            messagebox.showwarning("Validation", "La direction est obligatoire.")
            return

        data = (
            self.matricule_var.get().strip(),
            self.badge_var.get().strip() or None,
            self.nom_var.get().strip().upper(),
            self.prenom_var.get().strip().title(),
            self.type_person_var.get(),
            self.fonction_var.get().strip(),
            self.activite_var.get().strip() or None,
            self.division_var.get().strip() or None,
            self.direction_var.get().strip(),
            self.departement_var.get().strip() or None,
            self.service_var.get().strip() or None,
            self.equipe_var.get().strip() or None,
            self.date_embauche.get_date().strftime('%Y-%m-%d') if self.date_embauche.get() else None,
            self.date_naissance.get_date().strftime('%Y-%m-%d') if self.date_naissance.get() else None,
            self.adresse_var.get().strip() or None,
            self.telephone_var.get().strip() or None,
            self.email_var.get().strip() or None,
            photo, # Ajout du champ photo
            self.statut_var.get(),
            1 if self.concerne_var.get() else 0,
            self.type_quart_var.get(),
            self.heure_entree_var.get().strip() or '08:00:00',
            self.heure_sortie_var.get().strip() or '17:00:00',
            self.tolerance_entree_var.get(),
            self.tolerance_sortie_var.get()
        )

        try:

            if self.personnel_id:
                
             # Si pas de nouvelle photo, on garde l'ancienne
                if not (hasattr(self, 'photo_path') and self.photo_path):
                  photo = self.personnel['photo'] if self.personnel else None
                  
                self.app.db.update_personnel(self.personnel_id, data)
                messagebox.showinfo("Succès", "Personnel modifié avec succès.")
            else:
                self.app.db.add_personnel(data)
                messagebox.showinfo("Succès", "Personnel ajouté avec succès.")
            self.result = True
            self.dialog.destroy()
        except sqlite3.IntegrityError as e:
            if 'matricule' in str(e):
                messagebox.showerror("Erreur", "Ce matricule existe déjà.")
            elif 'badge_id' in str(e):
                messagebox.showerror("Erreur", "Ce badge ID est déjà utilisé.")
            else:
                messagebox.showerror("Erreur", f"Doublon détecté : {e}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'enregistrement : {e}")
            
            
            print("Nombre de champs :", len(data))
            if len(data) != 25:
               raise ValueError(f"Attendu 25, reçu {len(data)}")
            
# ------------------------------------------------------------
#   USER DIALOG (Ajout utilisateur)
# ------------------------------------------------------------
class UserDialog:
    def __init__(self, parent, app):
        self.app = app
        self.result = False
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Ajouter un utilisateur")
        self.dialog.geometry("400x450")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.create_widgets()
        self.dialog.wait_window()

    def create_widgets(self):
        tk.Label(self.dialog, text="AJOUTER UN UTILISATEUR", font=('Arial',14,'bold')).pack(pady=10)
        f = tk.Frame(self.dialog, bg='white')
        f.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        tk.Label(f, text="Nom d'utilisateur *:").grid(row=0, column=0, sticky='w', pady=5)
        self.username = tk.Entry(f, width=30)
        self.username.grid(row=0, column=1, pady=5, padx=10)

        tk.Label(f, text="Email *:").grid(row=1, column=0, sticky='w', pady=5)
        self.email = tk.Entry(f, width=30)
        self.email.grid(row=1, column=1, pady=5, padx=10)

        tk.Label(f, text="Mot de passe *:").grid(row=2, column=0, sticky='w', pady=5)
        self.password = tk.Entry(f, show='*', width=30)
        self.password.grid(row=2, column=1, pady=5, padx=10)

        tk.Label(f, text="Confirmer *:").grid(row=3, column=0, sticky='w', pady=5)
        self.confirm = tk.Entry(f, show='*', width=30)
        self.confirm.grid(row=3, column=1, pady=5, padx=10)

        tk.Label(f, text="Nom *:").grid(row=4, column=0, sticky='w', pady=5)
        self.nom = tk.Entry(f, width=30)
        self.nom.grid(row=4, column=1, pady=5, padx=10)

        tk.Label(f, text="Prénom *:").grid(row=5, column=0, sticky='w', pady=5)
        self.prenom = tk.Entry(f, width=30)
        self.prenom.grid(row=5, column=1, pady=5, padx=10)

        tk.Label(f, text="Rôle *:").grid(row=6, column=0, sticky='w', pady=5)
        self.role = ttk.Combobox(f, values=['admin','superviseur','utilisateur'], width=28, state='readonly')
        self.role.set('utilisateur')
        self.role.grid(row=6, column=1, pady=5, padx=10)

        btn_f = tk.Frame(self.dialog)
        btn_f.pack(pady=20)
        tk.Button(btn_f, text="Enregistrer", command=self.save,
                  bg='#2ecc71', fg='white', width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_f, text="Annuler", command=self.dialog.destroy,
                  bg='#e74c3c', fg='white', width=15).pack(side=tk.LEFT, padx=10)

    def save(self):
        if not all([self.username.get(), self.email.get(), self.password.get(),
                    self.nom.get(), self.prenom.get()]):
            messagebox.showwarning("Validation", "Tous les champs sont obligatoires")
            return
        if self.password.get() != self.confirm.get():
            messagebox.showwarning("Validation", "Les mots de passe ne correspondent pas")
            return
        try:
            self.app.db.create_user(
                self.username.get(),
                self.password.get(),
                self.nom.get().upper(),
                self.prenom.get().title(),
                self.email.get(),
                self.role.get()
            )
            messagebox.showinfo("Succès", "Utilisateur créé")
            self.result = True
            self.dialog.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Erreur", "Ce nom d'utilisateur existe déjà")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur: {e}")


# ------------------------------------------------------------
#   EDIT PROFILE DIALOG
# ------------------------------------------------------------
class EditProfileDialog:
    def __init__(self, parent, app):
        self.app = app
        self.result = False
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Modifier le profil")
        self.dialog.geometry("400x250")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.create_widgets()
        self.dialog.wait_window()

    def create_widgets(self):
        tk.Label(self.dialog, text="MODIFIER MON PROFIL", font=('Arial',14,'bold')).pack(pady=10)
        f = tk.Frame(self.dialog, bg='white')
        f.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        tk.Label(f, text="Nom *:").grid(row=0, column=0, sticky='w', pady=5)
        self.nom = tk.Entry(f, width=30)
        self.nom.grid(row=0, column=1, pady=5, padx=10)
        self.nom.insert(0, self.app.current_user['nom'])

        tk.Label(f, text="Prénom *:").grid(row=1, column=0, sticky='w', pady=5)
        self.prenom = tk.Entry(f, width=30)
        self.prenom.grid(row=1, column=1, pady=5, padx=10)
        self.prenom.insert(0, self.app.current_user['prenom'])

        tk.Label(f, text="Email *:").grid(row=2, column=0, sticky='w', pady=5)
        self.email = tk.Entry(f, width=30)
        self.email.grid(row=2, column=1, pady=5, padx=10)
        self.email.insert(0, self.app.current_user.get('email', ''))

        btn_f = tk.Frame(self.dialog)
        btn_f.pack(pady=20)
        tk.Button(btn_f, text="Enregistrer", command=self.save,
                  bg='#2ecc71', fg='white', width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_f, text="Annuler", command=self.dialog.destroy,
                  bg='#e74c3c', fg='white', width=15).pack(side=tk.LEFT, padx=10)

    def save(self):
        if not all([self.nom.get(), self.prenom.get(), self.email.get()]):
            messagebox.showwarning("Validation", "Tous les champs sont obligatoires")
            return
        try:
            self.app.db.update_user(
                self.app.current_user['id'],
                nom=self.nom.get().upper(),
                prenom=self.prenom.get().title(),
                email=self.email.get()
            )
            messagebox.showinfo("Succès", "Profil mis à jour")
            self.result = True
            self.dialog.destroy()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur: {e}")


# ------------------------------------------------------------
#   CHANGE PASSWORD DIALOG
# ------------------------------------------------------------
class ChangePasswordDialog:
    def __init__(self, parent, app):
        self.app = app
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Changer le mot de passe")
        self.dialog.geometry("400x250")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.create_widgets()
        self.dialog.wait_window()

    def create_widgets(self):
        tk.Label(self.dialog, text="CHANGER LE MOT DE PASSE", font=('Arial',14,'bold')).pack(pady=10)
        f = tk.Frame(self.dialog, bg='white')
        f.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        tk.Label(f, text="Mot de passe actuel *:").pack(anchor='w', pady=5)
        self.old = tk.Entry(f, show='*', width=30)
        self.old.pack(pady=5)

        tk.Label(f, text="Nouveau mot de passe *:").pack(anchor='w', pady=5)
        self.new = tk.Entry(f, show='*', width=30)
        self.new.pack(pady=5)

        tk.Label(f, text="Confirmer *:").pack(anchor='w', pady=5)
        self.confirm = tk.Entry(f, show='*', width=30)
        self.confirm.pack(pady=5)

        btn_f = tk.Frame(self.dialog)
        btn_f.pack(pady=20)
        tk.Button(btn_f, text="Changer", command=self.change,
                  bg='#2ecc71', fg='white', width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_f, text="Annuler", command=self.dialog.destroy,
                  bg='#e74c3c', fg='white', width=15).pack(side=tk.LEFT, padx=10)

    def change(self):
        old = self.old.get()
        new = self.new.get()
        confirm = self.confirm.get()
        if not old or not new or not confirm:
            messagebox.showwarning("Validation", "Tous les champs sont obligatoires")
            return
        if new != confirm:
            messagebox.showwarning("Validation", "Les mots de passe ne correspondent pas")
            return
        user = self.app.db.authenticate_user(self.app.current_user['username'], old)
        if not user:
            messagebox.showerror("Erreur", "Mot de passe actuel incorrect")
            return
        # Mettre à jour le mot de passe
        conn = None
        try:
            conn = sqlite3.connect(self.app.db.db_name, timeout=30)
            cursor = conn.cursor()
            hashed = hashlib.sha256(new.encode()).hexdigest()
            cursor.execute('UPDATE utilisateurs SET password=? WHERE id=?', (hashed, self.app.current_user['id']))
            conn.commit()
            messagebox.showinfo("Succès", "Mot de passe modifié")
            self.dialog.destroy()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur: {e}")
        finally:
            if conn: conn.close()


# ============================================================
#   CLASSE PRINCIPALE
# ============================================================
class PointageApp:
    def __init__(self):
        self.db = Database()
        self.current_user = None

    def run(self):
        self.show_login()

    def show_login(self):
        root = tk.Tk()
        LoginWindow(root, self)
        root.mainloop()

    def show_main_window(self):
        root = tk.Tk()
        MainWindow(root, self)
        root.mainloop()


# ============================================================
#   POINT D'ENTRÉE
# ============================================================
if __name__ == "__main__":
    try:
        print("="*60)
        print("🚀 SYSTÈME DE POINTAGE - VERSION FINALE (TOUT INTÉGRÉ)")
        print("="*60)
        # Dépendances
        pkgs = ['tkcalendar', 'matplotlib', 'pandas', 'pillow', 'openpyxl']
        import subprocess, sys
        for p in pkgs:
            try:
                __import__(p.replace('-','_'))
            except ImportError:
                print(f"📥 Installation de {p}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", p])
        app = PointageApp()
        app.run()
    except KeyboardInterrupt:
        print("\n👋 Arrêt.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Erreur: {e}")
        import traceback
        traceback.print_exc()
        input("\nAppuyez sur Entrée...")
        sys.exit(1)