# -*- coding: utf-8 -*-
import os
import re
import sys
import pandas as pd
import shutil
from datetime import datetime

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QTreeWidget, QTreeWidgetItem, QHeaderView,
                              QPushButton, QComboBox, QProgressBar, QMenu, QMessageBox,
                              QInputDialog, QFileDialog, QScrollArea, QFrame, QStyledItemDelegate)
from PySide6.QtCore import Qt, QSize, QPoint, QTimer, QRect
from PySide6.QtGui import (QColor, QBrush, QPalette, QAction, QClipboard, QGuiApplication,
                          QPainter, QLinearGradient, QIcon, QFont, QPen)

def ajouter_poids_of(df, fichier_recette):
    try:
        recette_df = pd.read_excel(fichier_recette, engine="openpyxl")
        recette_df.columns = [str(col).strip().upper() for col in recette_df.columns]

        if "CODE ARTICLE" not in recette_df.columns or "QUANTITÉ MATIÈRE (KG)" not in recette_df.columns:
            print("❌ Colonnes manquantes dans le fichier de recette.")
            df["Le poids d'OF"] = 0
            return df

        recette_agg = recette_df.groupby("CODE ARTICLE")["QUANTITÉ MATIÈRE (KG)"].sum().reset_index()
        recette_agg.rename(columns={"QUANTITÉ MATIÈRE (KG)": "poids_unitaire"}, inplace=True)

        df["ANCIENCODE"] = df["ANCIENCODE"].astype(str).str.strip().str.upper()
        recette_agg["CODE ARTICLE"] = recette_agg["CODE ARTICLE"].astype(str).str.strip().str.upper()

        df = df.merge(recette_agg, left_on="ANCIENCODE", right_on="CODE ARTICLE", how="left")

        df["QUANTITEOF"] = pd.to_numeric(df["QUANTITEOF"], errors="coerce").fillna(0)
        df["poids_unitaire"] = pd.to_numeric(df["poids_unitaire"], errors="coerce").fillna(0)
        df["Le poids d'OF"] = ((df["QUANTITEOF"] / 1000) * df["poids_unitaire"]).round(4)

        df.drop(columns=["CODE ARTICLE", "poids_unitaire"], errors="ignore", inplace=True)

    except Exception as e:
        print(f"⚠️ Erreur lors du calcul de 'Le poids d'OF': {e}")
        df["Le poids d'OF"] = 0

    return df

class TreeItemDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        painter.save()
        pen = QPen()
        pen.setColor(Qt.gray)
        pen.setWidth(1)
        painter.setPen(pen)
        painter.drawRect(option.rect.adjusted(0, 0, -1, -1))
        painter.restore()

COLORS = {
    'background': '#f8f9fa',
    'primary': '#4e73df',
    'primary_light': '#858796',
    'primary_dark': '#2e59d9',
    'secondary': '#6c757d',
    'accent': '#36b9cc',
    'text': '#3a3b45',
    'text_light': '#858796',
    'success': '#28a745',
    'warning': '#ffc107',
    'border': '#dddfeb',
    'error': '#dc3545',
    'header': '#5a5c69',
    'odd_row': '#ffffff',
    'even_row': '#f8f9fc',
    'modified_row': '#fff3e0',
    'highlight': '#ffcccc',
    'gradient_start': '#007bff',
    'gradient_end': '#003f7f',
    'produit_ce_jour': '#e6f7ff'
}

# Système de familles
FAMILLE_RULES = {
    'H07': 'Domestique',
    'H05': 'Domestique',
    'NYM': 'Domestique',
    'A05': 'Domestique',
    'A03': 'Domestique',
    '05VV-F': 'Domestique',
    'CABLE DE RACCORDEMENT': 'Torsade',
    'CR': 'Torsade',
    'FIL ROND': 'Fil guipé',
    'FIL PLAT': 'Fil guipé',
    'FIL DE BOBINAGE ROND': 'Fil guipé',
    'Cu':'Cuivre Nu',
    'CABLE COAXIAL': 'Television',
    'CÂBLE COAXIAL': 'Television',
    'CABLE COXIAL': 'Television',

    'Cu dur CL2 RC': 'Cuivre Nu',
    'Cu rácuit CL2 70 mm2': 'Cuivre Nu',
    'Cu rácuit': 'Cuivre Nu',
    'FILS EMAILLES': 'Cuivre Nu',
    'Fil de contact BC': 'Cuivre Nu',
   

    'AGS': 'AGS/alu acier',
    'Alu/Acier': 'AGS/alu acier',
    'H1Z2Z2-K': 'Solair',
    'N2XY': 'Industriel Cu',
    'NA2XH-O  (HS) 1X240 MM² 0.6/1 KV': 'MT HHFR',
    'NA2XS(F)Y': 'MT AL',
    'NA2XSBY': 'MT AL',
    
    'N2XSY': 'MT Cu',
    'N2XSBY': 'MT Cu',
    'NA2XH-O (HS) 1X120 1.8/3KV': 'MT HHFR',
    'NA2XH (HS) 1X400  MM2 1.8/3KV': 'MT HHFR',
    'NA2XH-O (HS) 1X300 MMŠ 1.8/3KV': 'MT HHFR',
    'NA2XH': ' HHFR AL',
    'N2XBY': 'Industriel Cu',
    'N2XRGY-J': 'Industriel Cu',
    
    'N2XBH': 'HHFR 1000 V',
    'N2XH': 'HHFR 1000 V',
    'RZ1': 'HHFR 1000 V',
    'NA2XY': 'Industriel AL',
    'N2X(F)Y': 'MT Cu',
    'N2XBH': 'HHFR 1000 V',
}

class FilterBar(QWidget):
    def __init__(self, parent=None, apply_filter_callback=None):
        super().__init__(parent)
        self.entries = {}
        self.apply_filter_callback = apply_filter_callback
        self.setup_ui()

    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        
        columns = ["DESIGNATION PRODUIT","N° OF", "OBS", "Etat d'OF", "", "ETAT", "Famille"]
        
        for col in columns:
            lbl = QLabel(col)
            lbl.setStyleSheet("font-family: Microsoft Sans Serif; font-size: 9pt;")
            layout.addWidget(lbl)
            
            entry = QLineEdit()
            entry.setMinimumWidth(200)
            entry.setStyleSheet("""
                QLineEdit {
                    font-family: Microsoft Sans Serif;
                    font-size: 11pt;
                    border: 1px solid #ccc;
                    padding: 2px;
                }
            """)
            entry.textChanged.connect(self.apply_filter_callback)
            
            entry.setContextMenuPolicy(Qt.CustomContextMenu)
            entry.customContextMenuRequested.connect(lambda pos, widget=entry: self.show_context_menu(widget, pos))
            
            layout.addWidget(entry)
            self.entries[col] = entry
        
        layout.addStretch()

    def get_filters(self):
        return {col: entry.text().strip().lower() for col, entry in self.entries.items() if entry.text().strip()}
    
    def show_context_menu(self, widget, pos):
        menu = QMenu(self)
        paste_action = QAction("Coller", self)
        paste_action.triggered.connect(lambda: widget.paste())
        menu.addAction(paste_action)
        menu.exec_(widget.mapToGlobal(pos))

class GradientHeader(QWidget):
    def __init__(self, parent=None, color1=COLORS['gradient_start'], color2=COLORS['gradient_end']):
        super().__init__(parent)
        self.color1 = QColor(color1)
        self.color2 = QColor(color2)
        self.setFixedHeight(40)
        
    def paintEvent(self, event):
        painter = QPainter(self)
        gradient = QLinearGradient(0, 0, self.width(), 0)
        gradient.setColorAt(0, self.color1)
        gradient.setColorAt(1, self.color2)
        painter.fillRect(self.rect(), gradient)
        
        painter.setPen(QColor(255, 255, 255))
        font = QFont("Segoe UI", 12, QFont.Bold)
        painter.setFont(font)
        painter.drawText(self.rect(), Qt.AlignCenter, "SYSTÈME DE SUIVI SGOF")
        painter.end()

class ModernButton(QPushButton):
    def __init__(self, text, icon=None, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QPushButton {
                font-family: 'Segoe UI', Arial;
                font-size: 11pt;
                font-weight: 500;
                color: white;
                background-color: %s;
                padding: 8px 12px;
                border: none;
                border-radius: 6px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: %s;
            }
            QPushButton:pressed {
                background-color: %s;
            }
        """ % (COLORS['primary'], COLORS['primary_dark'], COLORS['accent']))
        
        if icon:
            self.setIcon(icon)
            self.setIconSize(QSize(16, 16))

class FamilleManager(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gestion des Familles")
        self.resize(900, 600)
        self.familles = {}
        self.parent = parent
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Search frame
        search_frame = QWidget()
        search_layout = QHBoxLayout(search_frame)
        
        lbl_search = QLabel("Rechercher:")
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("Rechercher par N° OF, Designation ou Famille")
        self.search_entry.textChanged.connect(self.filter_familles)
        
        search_layout.addWidget(lbl_search)
        search_layout.addWidget(self.search_entry)
        
        # Tree widget
        self.tree = QTreeWidget()
        self.tree.setColumnCount(4)
        self.tree.setHeaderLabels(["N° OF", "DESIGNATION", "Famille Auto", "Famille Manuelle"])
        self.tree.setAlternatingRowColors(True)
        self.tree.setSelectionMode(QTreeWidget.SingleSelection)
        self.tree.header().setSectionResizeMode(QHeaderView.Interactive)
        
        # Info frame
        info_frame = QFrame()
        info_frame.setFrameShape(QFrame.StyledPanel)
        info_layout = QHBoxLayout(info_frame)
        
        self.of_label = QLabel("N° OF:")
        self.of_value = QLabel("")
        self.designation_label = QLabel("DESIGNATION:")
        self.designation_value = QLabel("")
        
        info_layout.addWidget(self.of_label)
        info_layout.addWidget(self.of_value)
        info_layout.addWidget(self.designation_label)
        info_layout.addWidget(self.designation_value)
        info_layout.addStretch()
        
        # Edit frame
        edit_frame = QFrame()
        edit_frame.setFrameShape(QFrame.StyledPanel)
        edit_layout = QVBoxLayout(edit_frame)
        
        lbl_famille = QLabel("Famille:")
        self.famille_entry = QLineEdit()
        self.famille_entry.setPlaceholderText("Entrez la famille")
        
        btn_load_auto = QPushButton("Charger Auto")
        btn_load_auto.clicked.connect(self.load_auto_famille)
        
        btn_save = QPushButton("Enregistrer")
        btn_save.clicked.connect(self.save_famille)
        
        btn_delete = QPushButton("Supprimer")
        btn_delete.clicked.connect(self.delete_famille)
        
        edit_layout.addWidget(lbl_famille)
        edit_layout.addWidget(self.famille_entry)
        edit_layout.addWidget(btn_load_auto)
        edit_layout.addWidget(btn_save)
        edit_layout.addWidget(btn_delete)
        
        # Main layout
        layout.addWidget(search_frame)
        layout.addWidget(self.tree)
        layout.addWidget(info_frame)
        layout.addWidget(edit_frame)
        
        # Connect signals
        self.tree.itemSelectionChanged.connect(self.update_info)
        
    def load_familles(self):
        """Charger les familles depuis le fichier"""
        familles_path = os.path.join(self.parent.data_folder, "familles.csv")
        if os.path.exists(familles_path):
            try:
                with open(familles_path, "r", encoding="utf-8") as f:
                    for line in f:
                        parts = line.strip().split(",")
                        if len(parts) >= 2:
                            of_num, famille = parts[0], ",".join(parts[1:])
                            self.familles[of_num] = famille
            except Exception as e:
                print(f"Erreur lors du chargement des familles: {e}")
    
    def save_familles(self):
        """Sauvegarder les familles dans le fichier"""
        familles_path = os.path.join(self.parent.data_folder, "familles.csv")
        try:
            with open(familles_path, "w", encoding="utf-8") as f:
                for of_num, famille in self.familles.items():
                    f.write(f"{of_num},{famille}\n")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de sauvegarder les familles: {e}")
    
    def determine_famille(self, designation):
        """Déterminer automatiquement la famille basée sur DESIGNATION PRODUIT"""
        if not isinstance(designation, str):
            return ""
        
        designation = designation.strip().upper()
        for key, famille in FAMILLE_RULES.items():
            if key.upper() in designation:
                return famille
        return ""
    
    def populate_tree(self):
        """Remplir l'arbre avec les données"""
        self.tree.clear()
        
        if not hasattr(self.parent, 'df') or self.parent.df.empty:
            return
        
        for _, row in self.parent.df.iterrows():
            of_num = str(row.get("N° OF", ""))
            designation = str(row.get("DESIGNATION PRODUIT", ""))
            auto_famille = self.determine_famille(designation)
            manual_famille = self.familles.get(of_num, "")
            
            item = QTreeWidgetItem()
            item.setText(0, of_num)
            item.setText(1, designation)
            item.setText(2, auto_famille)
            item.setText(3, manual_famille)
            self.tree.addTopLevelItem(item)
    
    def filter_familles(self):
        """Filtrer les familles selon la recherche"""
        search_text = self.search_entry.text().lower()
        
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            match = False
            for col in range(4):
                if search_text in item.text(col).lower():
                    match = True
                    break
            item.setHidden(not match)
    
    def update_info(self):
        """Mettre à jour les infos de l'OF sélectionné"""
        selected = self.tree.currentItem()
        if selected:
            self.of_value.setText(selected.text(0))
            self.designation_value.setText(selected.text(1))
    
    def load_auto_famille(self):
        """Charger la famille automatique"""
        selected = self.tree.currentItem()
        if selected:
            auto_famille = selected.text(2)
            if auto_famille:
                self.famille_entry.setText(auto_famille)
    
    def save_famille(self):
        """Sauvegarder la famille manuelle"""
        selected = self.tree.currentItem()
        if not selected:
            QMessageBox.warning(self, "Avertissement", "Veuillez sélectionner un OF")
            return
            
        of_num = selected.text(0)
        famille = self.famille_entry.text().strip()
        
        if famille:
            self.familles[of_num] = famille
            selected.setText(3, famille)
            self.save_familles()
            QMessageBox.information(self, "Succès", "Famille enregistrée avec succès")
            
            # Mettre à jour l'arbre principal si nécessaire
            if hasattr(self.parent, 'update_tree'):
                self.parent.update_tree()
    
    def delete_famille(self):
        """Supprimer la famille manuelle"""
        selected = self.tree.currentItem()
        if not selected:
            return
            
        of_num = selected.text(0)
        if of_num in self.familles:
            del self.familles[of_num]
            selected.setText(3, "")
            self.save_familles()
            self.famille_entry.clear()
            QMessageBox.information(self, "Succès", "Famille supprimée avec succès")
            
            # Mettre à jour l'arbre principal si nécessaire
            if hasattr(self.parent, 'update_tree'):
                self.parent.update_tree()

class ApplicationTraitementSGOF(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize directories first
        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.pdf_folder = os.path.join(self.base_dir, "pdfs")
        self.data_folder = os.path.join(self.base_dir, "data")
        
        # Create directories if they don't exist
        os.makedirs(self.pdf_folder, exist_ok=True)
        os.makedirs(self.data_folder, exist_ok=True)
        
        # Initialize variables
        self.df = pd.DataFrame()
        self.current_file_path = None
        self.poids_visible = True
        self.recette_data = None
        self.filtered_df = None
        self.sort_column = None
        self.sort_reverse = False
        self.filter_text = ""
        self.context_menu_item = None
        self.context_menu_column = None
        self.double_click_item = None
        self.double_click_column = None
        self.familles = {}  # Dictionnaire pour stocker les familles {N° OF: Famille}
        
        # Initialize UI
        self.setWindowTitle("🔷 Système de Suivi SGOF")
        self.resize(1200, 800)
        
        # Setup etat options
        self.etat_of_options = [
            " ", "Programmé H2", "Le Cord Nu existe H2", "Attente Isolation", 
            "Attente Assemblage", "Attente Gainage ext ", "Attente Gainage (Séparation)", 
            "Attente Recoouvrement", "Au Champ D'esser  ", "Cloturer(livré)","Clacage Attante Reboubinage ","L'OF Annuler"
        ]
        
        self.etat_colors = {
            "Programmé H2": "#d2b48c",
            "Le Cord Nu existe H2": "#FFFF00",
            "Attente Isolation": "#87CEEB",
            "Attente Assemblage": "#cce5ff",
            "Attente Gainage ext ": "#e2e3e5",
            "Attente Gainage (Séparation)": "#d1ecf1",
            "Attente Recoouvrement": "#FFC0CB",
            "Au Champ D'esser  ": "#28a745",
            "Cloturer(livré)": "#dc3545",
            "Clacage Attante Reboubinage ": "#FFD700",
            "L'OF Annuler": "#FFA500"
        }

        # Setup UI components
        self.setup_menu_bar()
        self.setup_ui()
        
        # Load data
        self.load_recette()
        self.load_familles()  # Charger les familles au démarrage
        self.load_latest_file()
        
        # Setup auto-save timer
        self.timer = QTimer()
        self.timer.timeout.connect(self.save_changes_auto)
        self.timer.start(20000)  # 20 seconds
        
        # Process data
        if not self.df.empty:
            self.traitement_automatique()
        if hasattr(self, 'df') and not self.df.empty:
            self.setup_column_visibility_menu()    

    def load_familles(self):
        """Charger les familles depuis le fichier"""
        familles_path = os.path.join(self.data_folder, "familles.csv")
        if os.path.exists(familles_path):
            try:
                with open(familles_path, "r", encoding="utf-8") as f:
                    for line in f:
                        parts = line.strip().split(",")
                        if len(parts) >= 2:
                            of_num, famille = parts[0], ",".join(parts[1:])
                            self.familles[of_num] = famille
            except Exception as e:
                print(f"Erreur lors du chargement des familles: {e}")

    def save_familles(self):
        """Sauvegarder les familles dans le fichier"""
        familles_path = os.path.join(self.data_folder, "familles.csv")
        try:
            with open(familles_path, "w", encoding="utf-8") as f:
                for of_num, famille in self.familles.items():
                    f.write(f"{of_num},{famille}\n")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de sauvegarder les familles: {e}")

    def determine_famille(self, designation):
        """Déterminer automatiquement la famille basée sur DESIGNATION PRODUIT"""
        if not isinstance(designation, str):
            return ""
        
        designation = designation.strip().upper()
        for key, famille in FAMILLE_RULES.items():
            if key.upper() in designation:
                return famille
        return ""
    def save_column_states(self):
        """حفظ حالة الأعمدة"""
        if not hasattr(self, 'tree'):
            return
        
        states = {}
        header = self.tree.header()
        for i in range(header.count()):
            col_name = header.model().headerData(i, Qt.Horizontal)
            states[col_name] = not self.tree.isColumnHidden(i)
    
        # حفظ states في ملف أو متغير
    def load_column_states(self):
        """استعادة حالة الأعمدة"""
        # ... قراءة الحالة المحفوظة ...
        for i in range(self.tree.header().count()):
            col_name = self.tree.header().model().headerData(i, Qt.Horizontal)
            if col_name in saved_states:
                self.tree.setColumnHidden(i, not saved_states[col_name])    
    def save_changes_auto(self):
        try:
            if hasattr(self, 'df') and not self.df.empty:
                sauvegarde_path = os.path.join(self.data_folder, "SGOF_sauvegarde_auto.xlsx")

                # التأكد من وجود عمود OBSERVATION
                if "OBSERVATION" not in self.df.columns:
                    self.df["OBSERVATION"] = ""

                colonnes_finales = [
                    "ANCIENCODE", "DESIGNATION PRODUIT", "N° OF", "DATEOF", "PREVUCLOTURE",
                    "QUANTITEOF", "TOLER.MIN %", "TOLER.MAX %", "OBS", "ETAT", "REALISE",
                    "Etat d'OF","Famille", "Le poids d'OF","OBSERVATION","HOLE"
                ]

                colonnes_existantes = [col for col in colonnes_finales if col in self.df.columns]
                self.df[colonnes_existantes].to_excel(sauvegarde_path, index=False)
                print("✅ Sauvegarde automatique avec colonnes choisies :", colonnes_existantes)
            else:
                print("ℹ️ Aucun DataFrame chargé, aucune sauvegarde effectuée.")
        except Exception as e:
            print(f"⚠️ Erreur sauvegarde auto: {e}")

    def setup_menu_bar(self):
        
        menu_bar = self.menuBar() 
        # قائمة Fichier (الملف)
        menu_fichier = menu_bar.addMenu("📁 Fichier")
        action_exporter = QAction("📤 Exporter les données", self)
        action_exporter.triggered.connect(self.exporter_donnees)
        menu_fichier.addAction(action_exporter)
        action_fermer = QAction("❌ Fermer", self)
        action_fermer.triggered.connect(self.close)
        menu_fichier.addAction(action_fermer)

        # قائمة Affichage (العرض)
        menu_affichage = menu_bar.addMenu("👁️ Affichage")
    
        # قائمة فرعية لإظهار/إخفاء الأعمدة
        self.menu_colonnes = menu_affichage.addMenu("Afficher les colonnes")
    
        # إضافة خيارات إظهار/إخفاء الأعمدة
        self.setup_column_visibility_menu()
    
        # باقي القوائم...
        menu_stats = menu_bar.addMenu("📊 Statistiques")
        action_famille = QAction("Totaux par famille", self)
        action_famille.triggered.connect(self.afficher_totaux_famille)
        menu_stats.addAction(action_famille)
    
        menu_aide = menu_bar.addMenu("❓ Aide")
        action_apropos = QAction("À propos", self)
        action_apropos.triggered.connect(self.afficher_apropos)
        menu_aide.addAction(action_apropos) 
    def toggle_column_visibility(self, column_name, visible):
        """تبديل إظهار/إخفاء العمود"""
        if not hasattr(self, 'tree'):
            return
    
        # البحث عن العمود في TreeWidget
        header = self.tree.header()
        for i in range(header.count()):
            if header.model().headerData(i, Qt.Horizontal) == column_name:
                self.tree.setColumnHidden(i, not visible)
                break
    def setup_column_visibility_menu(self):
        """تهيئة قائمة إظهار/إخفاء الأعمدة"""
        if not hasattr(self, 'df') or self.df.empty:
            return
    
        # مسح القائمة الحالية إن وجدت
        self.menu_colonnes.clear()
    
        # إضافة خيار لكل عمود
        for col in self.df.columns:
            action = QAction(col, self, checkable=True)
            action.setChecked(True)  # جميع الأعمدة ظاهرة افتراضياً
            action.triggered.connect(lambda checked, c=col: self.toggle_column_visibility(c, checked))
            self.menu_colonnes.addAction(action)       
    def load_recette(self):
        recette_path = os.path.join(self.data_folder, "recette 2025.xlsx")
        if os.path.exists(recette_path):
            try:
                self.recette_data = pd.read_excel(recette_path)
                print("✅ Recette chargée avec succès.")
            except Exception as e:
                print(f"⚠️ Erreur lors du chargement de la recette: {e}")
                self.recette_data = pd.DataFrame()
        else:
            print("⚠️ Fichier 'recette 2025.xlsx' non trouvé.")
            self.recette_data = pd.DataFrame()

    def exporter_donnees(self):
        self.exporter_resultats()
    
    def handle_combo_poids_change(self, value):
        if value == "Afficher le poids":
            self.poids_visible = True
            self.update_tree()
        elif value == "Masquer le poids":
            self.poids_visible = False
            self.update_tree()
        elif value == "Recalculer":
            self.traiter_donnees()
            self.update_tree()

    def load_latest_file(self):
        latest_file = None
        latest_mtime = 0
        
        if not os.path.exists(self.data_folder):
            os.makedirs(self.data_folder)
            return
            
        for file in os.listdir(self.data_folder):
            if file.endswith(".xlsx") and not file.startswith("~$"):
                full_path = os.path.join(self.data_folder, file)
                mtime = os.path.getmtime(full_path)
                if mtime > latest_mtime:
                    latest_mtime = mtime
                    latest_file = full_path

        if latest_file:
            try:
                self.df = pd.read_excel(latest_file)
                if "OBSERVATION" not in self.df.columns:
                    self.df["OBSERVATION"] = ""
                self.current_file_path = latest_file
                print(f"✅ Fichier chargé automatiquement: {os.path.basename(latest_file)}")
            except Exception as e:
                print(f"⚠️ Erreur de chargement du fichier: {e}")
                self.df = pd.DataFrame()
        else:
            print("⚠️ Aucun fichier Excel trouvé dans le dossier 'data'.")
            self.df = pd.DataFrame()

    def afficher_totaux_famille(self):
        if not hasattr(self, 'df') or self.df.empty:
            QMessageBox.warning(self, "Avertissement", "Aucune donnée chargée.")
            return

        # Vérifier si la colonne Famille existe
        if "Famille" not in self.df.columns:
            self.df["Famille"] = self.df["DESIGNATION PRODUIT"].apply(self.determine_famille)

        # Calculer les totaux
        total_familles = self.df.groupby("Famille")["Le poids d'OF"].sum().reset_index()
        total_familles.rename(columns={"Le poids d'OF": "Total (kg)"}, inplace=True)

        # Créer une fenêtre pour afficher les résultats
        result_window = QDialog(self)
        result_window.setWindowTitle("Totaux par famille")
        result_window.resize(400, 300)
        
        layout = QVBoxLayout()
        
        # Tableau des résultats
        table = QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["Famille", "Total (kg)"])
        table.setRowCount(len(total_familles))
        
        for i, (_, row) in enumerate(total_familles.iterrows()):
            table.setItem(i, 0, QTableWidgetItem(row["Famille"]))
            table.setItem(i, 1, QTableWidgetItem(f"{row['Total (kg)']:.2f}"))
        
        table.resizeColumnsToContents()
        
        # Bouton d'export
        btn_export = QPushButton("Exporter vers Excel")
        btn_export.clicked.connect(lambda: self.export_totaux_famille(total_familles))
        
        layout.addWidget(table)
        layout.addWidget(btn_export)
        result_window.setLayout(layout)
        
        result_window.exec_()

    def export_totaux_famille(self, total_familles):
        today_str = datetime.now().strftime("%Y-%m-%d")
        export_path = os.path.join(self.data_folder, f"totaux_famille_{today_str}.xlsx")
        
        try:
            total_familles.to_excel(export_path, index=False)
            QMessageBox.information(self, "Succès", f"Fichier exporté :\n{export_path}")
            os.startfile(self.data_folder)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'export :\n{str(e)}")

    def afficher_par_jour(self):
        QMessageBox.information(self, "Statistiques", "✅ Production par jour en cours de calcul...")

    def afficher_apropos(self):
        QMessageBox.information(self, "À propos", "Système de Suivi de Production v1.0\nDéveloppé par nour.")

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
    
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
    
        self.gradient_header = GradientHeader()
        main_layout.addWidget(self.gradient_header)
    
        self.filter_bar = FilterBar(self, self.apply_filters)
        main_layout.addWidget(self.filter_bar)
    
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet(f"color: {COLORS['border']};")
        main_layout.addWidget(separator)
    
        self.setup_toolbar()
        main_layout.addWidget(self.toolbar)
    
        tree_container = QWidget()
        tree_container.setStyleSheet(f"background-color: {COLORS['background']};")
        tree_layout = QVBoxLayout(tree_container)
        tree_layout.setContentsMargins(10, 10, 10, 10)
    
        self.setup_main_tree()
        tree_layout.addWidget(self.tree_frame)
        main_layout.addWidget(tree_container, 1)
    
        self.setup_status_bar()
        main_layout.addWidget(self.status_frame)
    
        self.etat_of_combobox = QComboBox()
        self.etat_of_combobox.addItems(self.etat_of_options)
        self.etat_of_combobox.setStyleSheet("""
            QComboBox {
                font-size: 9pt;
                padding: 1px 3px;
                margin: 0px;
                border: 1px solid #c0c0c0;
                min-width: 70px;
                max-width: 90px;
                height: 20px;
            }
            QComboBox::drop-down {
                width: 12px;
            }
        """)
           
        self.etat_of_combobox.hide()
        self.etat_of_combobox.activated.connect(self.on_etat_selected)
        self.etat_of_combobox.setItemDelegate(QStyledItemDelegate())
        self.etat_of_combobox.view().setFixedWidth(150)
        self.etat_of_combobox.view().setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        self.tree_menu = QMenu(self)
        self.tree_menu.setStyleSheet("""
            QMenu {
                font-family: 'Segoe UI', Arial;
                font-size: 10pt;
                background-color: white;
                border: 1px solid %s;
            }
            QMenu::item:selected {
               background-color: %s;
                color: white;
            }
        """ % (COLORS['border'], COLORS['primary']))
    
        copy_action = QAction(QIcon.fromTheme("edit-copy"), "Copier", self)
        copy_action.triggered.connect(self.copy_tree_cell)
        self.tree_menu.addAction(copy_action)

        # Ajouter l'action pour gérer les familles
        famille_action = QAction("Gérer les familles", self)
        famille_action.triggered.connect(self.gerer_familles)
        self.tree_menu.addAction(famille_action)

    def gerer_familles(self):
        """Ouvrir la fenêtre de gestion des familles"""
        manager = FamilleManager(self)
        manager.load_familles()
        manager.populate_tree()
        manager.exec_()

    def setup_toolbar(self):
        self.toolbar = QWidget()
        self.toolbar.setStyleSheet(f"background-color: {COLORS['background']};")
        toolbar_layout = QHBoxLayout(self.toolbar)
        toolbar_layout.setContentsMargins(10, 5, 10, 5)
        toolbar_layout.setSpacing(10)
    
        btn_actualiser = ModernButton("🔄 Actualiser")
        btn_actualiser.clicked.connect(self.traitement_automatique)
        toolbar_layout.addWidget(btn_actualiser)
    
        btn_exporter = ModernButton("📤 Exporter")
        btn_exporter.clicked.connect(self.exporter_resultats)
        toolbar_layout.addWidget(btn_exporter)
    
        btn_ajouter_etat = ModernButton("➕ Ajouter État")
        btn_ajouter_etat.clicked.connect(self.ajouter_etat)
        toolbar_layout.addWidget(btn_ajouter_etat)

        btn_familles = ModernButton("👨‍👩‍👧‍👦 Familles")
        btn_familles.clicked.connect(self.gerer_familles)
        toolbar_layout.addWidget(btn_familles)
    
        self.filter_entry = QLineEdit()
        self.filter_entry.setPlaceholderText("Rechercher...")
        self.filter_entry.setStyleSheet("""
            QLineEdit {
                font-family: 'Segoe UI', Arial;
                font-size: 10pt;
                padding: 8px;
                border: 1px solid %s;
                border-radius: 4px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 1px solid %s;
            }
        """ % (COLORS['border'], COLORS['primary']))
        self.filter_entry.textChanged.connect(self.apply_filters)
        toolbar_layout.addWidget(self.filter_entry)
    
        toolbar_layout.addStretch()

    def get_button_style(self):
        return """
            QPushButton {
                font-family: Microsoft Sans Serif;
                font-size: 10pt;
                font-weight: bold;
                color: white;
                background-color: %s;
                padding: 6px;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: %s;
            }
        """ % (COLORS['primary'], COLORS['primary_dark'])

    def setup_main_tree(self):
        self.tree_frame = QFrame()
        self.tree_frame.setFrameShape(QFrame.StyledPanel)
        self.tree_frame.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
            }}
        """)
    
        tree_layout = QVBoxLayout(self.tree_frame)
        tree_layout.setContentsMargins(0, 0, 0, 0)
    
        self.tree = QTreeWidget()
        self.tree.setItemDelegate(TreeItemDelegate(self.tree))
        self.tree.setStyleSheet("""
            QTreeWidget {
                gridline-color: #dee2e6;
                show-decoration-selected: 1;
                outline: 0;
                font-family: 'Segoe UI', Arial;
                font-size: 10pt;
                alternate-background-color: %s;
                border: none;
            }
            QHeaderView::section {
                background-color: %s;
                color: white;
                font-weight: bold;
                padding: 10px;
                border: none;
                font-size: 10pt;
            }
            QTreeWidget::item {
                padding: 5px 0;
            }
            QTreeWidget::item:hover {
                background-color: %s;
            }
        """ % (COLORS['even_row'], COLORS['header'], COLORS['highlight']))
    
        self.tree.setAlternatingRowColors(True)
        self.tree.setSelectionMode(QTreeWidget.SingleSelection)
        self.tree.setSelectionBehavior(QTreeWidget.SelectRows)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_tree_menu)
        self.tree.header().setDefaultAlignment(Qt.AlignCenter)
        self.tree.header().setSectionResizeMode(QHeaderView.Interactive)
        self.tree.header().setSectionsClickable(True)
        self.tree.header().sectionClicked.connect(self.on_header_click)
        self.tree.itemClicked.connect(self.on_tree_click)
        self.tree.itemDoubleClicked.connect(self.on_tree_double_click)
        
        tree_layout.addWidget(self.tree)

    def setup_status_bar(self):
        self.status_frame = QWidget()
        self.status_frame.setStyleSheet(f"""
            background-color: {COLORS['background']};
            border-top: 1px solid {COLORS['border']};
            padding: 5px;
        """)
    
        status_layout = QHBoxLayout(self.status_frame)
        status_layout.setContentsMargins(10, 5, 10, 5)
    
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid %s;
                border-radius: 3px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: %s;
            }
        """ % (COLORS['border'], COLORS['primary']))
        status_layout.addWidget(self.progress)
    
        self.status_label = QLabel("Prêt")
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', Arial;
            font-size: 9pt;
            color: %s;
        """ % COLORS['text_light'])
        status_layout.addWidget(self.status_label)
    
        version_label = QLabel("Version 2.0 © 2025 ETAT D'OF ELHANI")
        version_label.setStyleSheet("""
            font-family: 'Segoe UI', Arial;
            font-size: 9pt;
            color: %s;
        """ % COLORS['text_light'])
        status_layout.addWidget(version_label)

    def traitement_automatique(self):
        try:
            self.maj_statut("Recherche des fichiers SGOF...")
            self.progress.setValue(10)

            if not os.path.exists(self.data_folder):
                os.makedirs(self.data_folder)
                QMessageBox.information(self, "Information", 
                    f"Le dossier 'data' a été créé à :\n{self.data_folder}")
                self.maj_statut("Dossier data créé - Ajoutez vos fichiers SGOF")
                return

            fichiers_excel = [
                f for f in os.listdir(self.data_folder)
                if f.lower().endswith(('.xlsx', '.xls'))
            ]

            if not fichiers_excel:
                QMessageBox.information(self, "Information",
                    f"Aucun fichier Excel trouvé dans le dossier 'data'.\n\n"
                    f"Veuillez placer vos fichiers SGOF dans :\n{self.data_folder}")
                self.maj_statut("Aucun fichier Excel trouvé")
                return

            fichiers_utiles = [
                f for f in fichiers_excel
                if "sauvegarde_auto" not in f.lower() and "sauvegarde_fermeture" not in f.lower()
            ]

            if not fichiers_utiles:
                QMessageBox.information(self, "Information",
                    f"Aucun fichier SGOF trouvé (hors sauvegarde) dans :\n{self.data_folder}")
                self.maj_statut("Aucun fichier SGOF à traiter")
                return

            fichiers_tries = sorted(
                [(f, os.path.getmtime(os.path.join(self.data_folder, f))) for f in fichiers_utiles],
                key=lambda x: x[1],
                reverse=True
            )

            dernier_fichier = fichiers_tries[0][0]
            chemin_fichier = os.path.join(self.data_folder, dernier_fichier)

            self.maj_statut(f"Chargement du fichier : {dernier_fichier}")
            self.progress.setValue(50)

            self.traiter_fichier(chemin_fichier)

        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors du chargement automatique :\n{str(e)}")
            self.maj_statut("Échec du chargement automatique")
        finally:
            self.progress.setValue(100)

    def reordonner_colonnes(self):
        colonnes_finales = [
             "ANCIENCODE", "DESIGNATION PRODUIT", "N° OF","DATEOF","PREVUCLOTURE","QUANTITEOF"
             ,"OBS","ETAT","REALISE", "Etat d'OF","Famille", "Le poids d'OF","OBSERVATION","HOLE"
        ]

        # الاحتفاظ فقط بالأعمدة الموجودة فعلاً في DataFrame
        colonnes_presentes = [col for col in colonnes_finales if col in self.df.columns]

        # إضافة الأعمدة الأخرى (غير المتوقعة) في نهاية الجدول
        autres_colonnes = [col for col in self.df.columns if col not in colonnes_finales]

        # ترتيب الأعمدة النهائي
        self.df = self.df[colonnes_presentes + autres_colonnes]
        
    def traiter_fichier(self, chemin_fichier):
        try:
            self.df = None
            sauvegarde_path = os.path.join(self.data_folder, "SGOF_sauvegarde_auto.xlsx")

            df_nouveau = pd.read_excel(chemin_fichier)
            df_nouveau.columns = [col.strip().upper() for col in df_nouveau.columns]
            self.colonnes_sgof_origine = df_nouveau.columns.tolist()
            df_nouveau.rename(columns={"N° OF": "N° OF"}, inplace=True)

            # التأكد من وجود عمود OBSERVATION
            if "OBSERVATION" not in df_nouveau.columns:
                df_nouveau["OBSERVATION"] = ""

            if "ETAT" in df_nouveau.columns:
                df_nouveau["ETAT"] = df_nouveau["ETAT"].astype(str).str.strip().str.upper()
                df_nouveau.loc[df_nouveau["ETAT"].isin(["C", "S"]), "Etat d'OF"] = "Cloturer(livré)"
            if "Etat d'OF" not in df_nouveau.columns:
                df_nouveau["Etat d'OF"] = ""

            # Ajouter la colonne Famille si elle n'existe pas
            if "Famille" not in df_nouveau.columns:
                df_nouveau["Famille"] = df_nouveau["DESIGNATION PRODUIT"].apply(self.determine_famille)

            if os.path.exists(sauvegarde_path):
                df_ancien = pd.read_excel(sauvegarde_path)
                if "OBSERVATION" not in df_ancien.columns:
                    df_ancien["OBSERVATION"] = ""
            else:
                df_ancien = pd.DataFrame()

            if not df_ancien.empty and "Etat d'OF" in df_ancien.columns and "N° OF" in df_ancien.columns:
                df_etats = df_ancien[["N° OF", "Etat d'OF", "OBSERVATION", "Famille"]].copy()
                df_etats["N° OF"] = df_etats["N° OF"].astype(str)
                df_nouveau["N° OF"] = df_nouveau["N° OF"].astype(str)

                df_nouveau = df_nouveau.merge(df_etats, on="N° OF", how="left", suffixes=('', '_ancien'))

                if "Etat d'OF_ancien" in df_nouveau.columns:
                    df_nouveau["Etat d'OF"] = df_nouveau["Etat d'OF"].fillna(df_nouveau["Etat d'OF_ancien"])
                    df_nouveau.drop(columns=["Etat d'OF_ancien"], inplace=True)

                if "OBSERVATION_ancien" in df_nouveau.columns:
                    df_nouveau["OBSERVATION"] = df_nouveau["OBSERVATION"].astype(str)
                    df_nouveau["OBSERVATION"] = df_nouveau["OBSERVATION"].where(
                        df_nouveau["OBSERVATION"].str.strip() != "",
                        df_nouveau["OBSERVATION_ancien"]
                    )
                    df_nouveau.drop(columns=["OBSERVATION_ancien"], inplace=True)

                if "Famille_ancien" in df_nouveau.columns:
                    df_nouveau["Famille"] = df_nouveau["Famille"].fillna(df_nouveau["Famille_ancien"])
                    df_nouveau.drop(columns=["Famille_ancien"], inplace=True)

            if not df_ancien.empty and "N° OF" in df_ancien.columns:
                anciens_conservés = df_ancien[~df_ancien["N° OF"].astype(str).isin(df_nouveau["N° OF"].astype(str))]
                anciens_conservés = anciens_conservés.dropna(axis=1, how='all')
                df_nouveau = df_nouveau.dropna(axis=1, how='all')
                anciens_conservés = anciens_conservés.reindex(columns=df_nouveau.columns)

                if not anciens_conservés.empty:
                    anciens_conservés = anciens_conservés.dropna(axis=1, how='all')
                    anciens_conservés = anciens_conservés.reindex(columns=df_nouveau.columns)
                    df_final = pd.concat([anciens_conservés, df_nouveau], ignore_index=True)
                    df_final['N° OF'] = df_final['N° OF'].astype(str).str.strip()
                    df_final.drop_duplicates(subset='N° OF', keep='last', inplace=True)
                else:
                    df_final = df_nouveau.copy()
            else:
                df_final = df_nouveau.copy()

            self.df = ajouter_poids_of(df_final.copy(), os.path.join(self.data_folder, "recette 2025.xlsx"))

            self.df['N° OF'] = self.df['N° OF'].astype(str).str.strip()
            self.df.drop_duplicates(subset='N° OF', keep='last', inplace=True)
            tolerance = 0.05  # 5%
            condition_cloture = self.df["REALISE"] >= (self.df["QUANTITEOF"] * (1 - tolerance))
            self.df.loc[condition_cloture, "Etat d'OF"] = "Cloturer(livré)"
            self.df["REALISE"] = pd.to_numeric(self.df["REALISE"], errors="coerce").fillna(0)
            self.df["QUANTITEOF"] = pd.to_numeric(self.df["QUANTITEOF"], errors="coerce").fillna(0)
            

            self.df["ETAT"] = self.df["ETAT"].astype(str).str.upper()
            self.df.loc[self.df["ETAT"] == "A", "Etat d'OF"] = "L'OF Annuler"
            self.df = self.df.fillna("")
            self.df["Etat d'OF"] = self.df["Etat d'OF"].fillna("")

            if "OBSERVATION" not in self.df.columns:
                self.df["OBSERVATION"] = ""

            # Mettre à jour les familles si nécessaire
            if "Famille" not in self.df.columns:
                self.df["Famille"] = self.df["DESIGNATION PRODUIT"].apply(self.determine_famille)

            self.traiter_donnees()
            self.update_tree()
            self.maj_statut("Traitement réussi")
            self.progress.setValue(100)
            self.reordonner_colonnes()  # <-- أضف هذا قبل السطر التال
            self.df.to_excel(sauvegarde_path, index=False)

        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors du traitement :\n{str(e)}")
            self.maj_statut("Échec du traitement")

    def traiter_donnees(self):
        try:
            colonnes_a_supprimer = ['QUANTITELABO', 'TOLER.MIN %', 'TOLER.MAX %']
            for col in colonnes_a_supprimer:
                if col in self.df.columns:
                    self.df.drop(col, axis=1, inplace=True)
    
            colonnes_date = ['DATEOF', 'DATE_PREVUE', 'PREVUCLOTURE']
            for col in colonnes_date:
                if col in self.df.columns:
                    try:
                        self.df[col] = pd.to_datetime(self.df[col], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
                    except:
                        continue
        
            self.df['HOLE'] = self.df.apply(self.classifier_ligne, axis=1)
        
            if "Etat d'OF" not in self.df.columns:
                self.df["Etat d'OF"] = ""
        
            if "OBSERVATION" not in self.df.columns:
                self.df["OBSERVATION"] = ""
        
            # Mettre à jour les familles si nécessaire
            if "Famille" not in self.df.columns:
                self.df["Famille"] = self.df["DESIGNATION PRODUIT"].apply(self.determine_famille)
        
        except Exception as e:
            raise e
    
    def classifier_ligne(self, ligne):
        try:
            designation = str(ligne['DESIGNATION PRODUIT']).upper()
            
            types_speciaux = [
                r'CU\s*ÁTAMÁ\s*CL2',
                r'CU\s*RÁCUIT\s*CL2',
                r'CU\s*CL2\s*RC',
                r'CU\s*DUR\s*CL2',
                r'AGS',
                r'CUIVRE\s*ROND\s*RECUIT'
            ]
            
            for motif in types_speciaux:
                if re.search(motif, designation):
                    return "H2"
            
            section = self.extraire_section_cable(designation)
            if section is not None:
                return "H3" if section <= 10 else "H6"
            
            return "Non spécifié"
        except:
            return "Non spécifié"
    
    def extraire_section_cable(self, designation):
        if pd.isna(designation):
            return None
            
        designation = str(designation).upper()
        
        motifs = [
            r'H07[VZ][\- ][KR]\s*(\d+[,.]?\d*)',
            r'(\d+)\s*[Xx*G]\s*(\d+[,.]?\d*)',
            r'[Xx*G]\s*(\d+[,.]?\d*)',
            r'(\d+)\s*MM\s*[/\\]\s*(\d+)',
            r'(\d+)\s*[,.]\s*(\d+)\s*MM',
        ]
        
        for motif in motifs:
            correspondance = re.search(motif, designation)
            if correspondance:
                try:
                    section = correspondance.group(correspondance.lastindex).replace(',', '.')
                    return float(section)
                except:
                    continue
        return None
    
    def update_tree(self):
        self.tree.clear()
    
        if not hasattr(self, 'df') or self.df.empty:
            return
        if "OBSERVATION" not in self.df.columns:
            self.df["OBSERVATION"] = ""

        # Define column order
        colonnes_ordre = [
             "ANCIENCODE", "DESIGNATION PRODUIT", "N° OF","DATEOF","PREVUCLOTURE","QUANTITEOF",
             "OBS","ETAT","REALISE", "Etat d'OF","Famille", "Le poids d'OF","OBSERVATION","HOLE"
        ]
        self.tree.setColumnCount(len(self.df.columns))
        self.tree.setHeaderLabels(self.df.columns.tolist())
     
        # ملء البيانات
        for _, row in self.df.iterrows():
            item = QTreeWidgetItem()
            for i, col in enumerate(self.df.columns):
                item.setText(i, str(row[col]))
            self.tree.addTopLevelItem(item)
    
        # تحديث قائمة إظهار/إخفاء الأعمدة
        if hasattr(self, 'menu_colonnes'):
            self.setup_column_visibility_menu()
        # Add other columns not in the list
        colonnes_supplementaires = [col for col in self.df.columns if col not in colonnes_ordre]
        colonnes_affichees = colonnes_ordre + colonnes_supplementaires
        
        self.tree.setHeaderLabels(colonnes_affichees)
        self.tree.setColumnCount(len(colonnes_affichees))
    
        # Use filtered_df if it exists, otherwise use the full df
        display_df = self.filtered_df if hasattr(self, 'filtered_df') and self.filtered_df is not None else self.df
    
        if self.filter_text:
            global_mask = display_df.apply(
                lambda row: row.fillna('').astype(str).str.lower().str.strip().str.contains(
                    self.filter_text.lower(), regex=False
                ).any(), axis=1
            )
            display_df = display_df[global_mask]
    
        if hasattr(self.filter_bar, 'get_filters'):
            column_filters = self.filter_bar.get_filters()
        
            if column_filters:
                mask = pd.Series([True] * len(display_df))
            
                for col, filter_text in column_filters.items():
                    if col in display_df.columns:
                        col_mask = display_df[col].fillna('').astype(str).str.lower().str.strip().str.contains(
                            filter_text.lower(), regex=False
                        )
                        mask = mask & col_mask
            
                display_df = display_df[mask]
    
        for i, row in display_df.iterrows():
            item = QTreeWidgetItem()
            for j, col in enumerate(colonnes_affichees):
                item.setText(j, str(row[col]))
            
                if col == "Etat d'OF":
                    etat = str(row[col]).strip()
                    color = self.etat_colors.get(etat, COLORS['background'])
                    item.setBackground(j, QBrush(QColor(color)))
        
            self.tree.addTopLevelItem(item)
    
        for i in range(self.tree.columnCount()):
            self.tree.resizeColumnToContents(i)

    def apply_filters(self):
        try:
            self.filter_text = self.filter_entry.text().strip().lower()
            self.update_tree()
        except Exception as e:
            QMessageBox.critical(self, "Erreur de filtrage", f"Erreur lors du filtrage :\n{str(e)}")
    
    def on_header_click(self, logical_index):
        col_name = self.tree.headerItem().text(logical_index)
        
        if self.sort_column == col_name:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col_name
            self.sort_reverse = False
        
        self.sort_tree_data(col_name, self.sort_reverse)
    
    def sort_tree_data(self, column, reverse=False):
        if not hasattr(self, 'df') or self.df.empty:
            return

        try:
            if column == "N° OF":
                self.df["N° OF"] = self.df["N° OF"].astype(float)
                self.df = self.df.sort_values(by="N° OF", ascending=not reverse)
                self.df["N° OF"] = self.df["N° OF"].astype(int).astype(str)
            else:
                self.df = self.df.sort_values(by=column, key=lambda x: x.astype(str).str.lower(), ascending=not reverse)

            self.update_tree()

        except Exception as e:
            QMessageBox.critical(self, "Erreur de tri", f"Impossible de trier par {column}:\n{str(e)}")
    
    def on_tree_click(self, item, column):
        col_name = self.tree.headerItem().text(column)

        if col_name == "ANCIENCODE":
            code = item.text(column).strip()
            if code:
                pdf_path = os.path.join(self.pdf_folder, f"{code}.pdf")
                if os.path.isfile(pdf_path):
                    os.startfile(pdf_path)
                else:
                    QMessageBox.warning(
                        self,
                        "Fichier introuvable",
                        f"❌ Aucun fichier '{code}.pdf' dans le dossier 'pdfs'"
                    )

        if col_name == "Etat d'OF":
            self.selected_item = item
            self.selected_column = column

            rect = self.tree.visualItemRect(item)
            header = self.tree.header()
            column_pos = 0
            for i in range(column):
                column_pos += header.sectionSize(i)
    
            tree_global_pos = self.tree.mapToGlobal(QPoint(0, 0))
            column_width = header.sectionSize(column)
            self.etat_of_combobox.setFixedWidth(column_width)
    
            popup_x = tree_global_pos.x() + column_pos
            popup_y = tree_global_pos.y() + rect.y() + rect.height()
    
            screen_geometry = QGuiApplication.primaryScreen().availableGeometry()
            if popup_y + self.etat_of_combobox.height() > screen_geometry.bottom():
                popup_y = tree_global_pos.y() + rect.y() - self.etat_of_combobox.height()
    
            current_text = item.text(column)
            self.etat_of_combobox.setCurrentText(current_text)
    
            self.etat_of_combobox.setParent(self.tree.viewport())
            self.etat_of_combobox.move(popup_x - tree_global_pos.x(), popup_y - tree_global_pos.y())
            self.etat_of_combobox.show()
            self.etat_of_combobox.showPopup()

    def on_etat_selected(self, index):
        if not hasattr(self, 'selected_item') or not hasattr(self, 'selected_column'):
            print("[DEBUG] Aucune cellule sélectionnée. L'événement est ignoré.")
            return

        text = self.etat_of_combobox.itemText(index)
        new_value = text
        self.selected_item.setText(self.selected_column, new_value)

        color = self.etat_colors.get(new_value, COLORS['background'])
        self.selected_item.setBackground(self.selected_column, QBrush(QColor(color)))

        if hasattr(self, 'df'):
            try:
                n_of_value = None
                for i in range(self.tree.columnCount()):
                    if self.tree.headerItem().text(i).strip() == "N° OF":
                        n_of_value = self.selected_item.text(i).strip()
                        break

                if not n_of_value:
                    print("❌ Impossible de trouver la valeur N° OF depuis l'élément sélectionné.")
                    return

                print(f"[DEBUG] N° OF détecté dynamiquement: '{n_of_value}'")
                self.df["N° OF"] = self.df["N° OF"].astype(str).str.strip()

                n_of_value = n_of_value.upper()
                self.df["N° OF"] = self.df["N° OF"].str.upper()

                if n_of_value in self.df["N° OF"].values:
                    self.df.loc[self.df["N° OF"] == n_of_value, "Etat d'OF"] = new_value
                    print(f"[DEBUG] ✅ Mise à jour: N° OF = '{n_of_value}' → Etat d'OF = '{new_value}'")
                    ligne = self.df[self.df["N° OF"] == n_of_value]
                    print("[DEBUG] Donnée modifiée dans DF:")
                    print(ligne[["N° OF", "Etat d'OF"]])
                else:
                    print(f"❌ N° OF '{n_of_value}' non trouvé dans le DataFrame.")
                    print("[DEBUG] Liste des N° OF présents dans le DataFrame:")
                    print(self.df["N° OF"].unique())

                sauvegarde_path = os.path.join(self.data_folder, "SGOF_sauvegarde_auto.xlsx")
                self.df.to_excel(sauvegarde_path, index=False)
                print(f"[DEBUG] Données sauvegardées vers: {sauvegarde_path}")

            except Exception as e:
                print("Erreur mise à jour DataFrame:", e)

        self.etat_of_combobox.hide()
        if hasattr(self, 'selected_item'): del self.selected_item
        if hasattr(self, 'selected_column'): del self.selected_column

    def on_tree_double_click(self, item, column):
        col_name = self.tree.headerItem().text(column)
    
        if col_name.upper() == "OBSERVATION":
            self.double_click_item = item
            self.double_click_column = column
        
            current_text = item.text(column)
            new_text, ok = QInputDialog.getText(
                self,
                "Modifier OBSERVATION",
                "Entrez votre OBSERVATION:",
                QLineEdit.Normal,
                current_text
            )
        
            if ok and new_text is not None:
                item.setText(column, new_text)
            
                # Update DataFrame
                n_of_value = None
                for i in range(self.tree.columnCount()):
                    if self.tree.headerItem().text(i).strip().upper() == "N° OF":
                        n_of_value = item.text(i).strip()
                        break
            
                if n_of_value and hasattr(self, 'df'):
                    try:
                        self.df["N° OF"] = self.df["N° OF"].astype(str).str.strip()
                        self.df.loc[self.df["N° OF"] == n_of_value, "OBSERVATION"] = new_text
                        print(f"[DEBUG] ✅ OBSERVATION mise à jour pour N° OF = '{n_of_value}'")
                        save_path_auto = os.path.join(self.data_folder, "SGOF_sauvegarde_auto.xlsx")
                        save_path_fermeture = os.path.join(self.data_folder, "SGOF_sauvegarde_fermeture.xlsx")

                        self.df.to_excel(save_path_auto, index=False)
                        self.df.to_excel(save_path_fermeture, index=False)

                        print(f"[DEBUG] 💾 OBSERVATION sauvegardée dans les deux fichiers.") 
                    except Exception as e:
                        print(f"⚠️ Erreur lors de la mise à jour de l'OBSERVATION: {e}")
    
    def exporter_resultats(self):
        try:
            aujourdhui = datetime.now().strftime("%d-%m-%Y")
            fichier_sortie = os.path.join(self.data_folder, f"SGOF_Traite_{aujourdhui}.xlsx")

            # التأكد من وجود عمود OBSERVATION
            if "OBSERVATION" not in self.df.columns:
                self.df["OBSERVATION"] = ""

            colonnes_ordre = [
                "ANCIENCODE", "DESIGNATION PRODUIT", "N° OF",
                "REALISE", "QUANTITEOF", "Etat d'OF","Famille", "Le poids d'OF","OBSERVATION","HOLE"
            ]

            colonnes_supplementaires = [col for col in self.df.columns if col not in colonnes_ordre]
            colonnes_export = colonnes_ordre + colonnes_supplementaires

            self.df[colonnes_export].to_excel(fichier_sortie, index=False)
            QMessageBox.information(self, "Succès", f"Résultats exportés vers :\n{fichier_sortie}")
            os.startfile(self.data_folder)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'export :\n{str(e)}")
    
    def ajouter_etat(self):
        new_etat, ok = QInputDialog.getText(self, "Ajouter un État", "Entrez le nouvel état:")
        if ok and new_etat and new_etat not in self.etat_of_options:
            self.etat_of_options.append(new_etat)
            self.etat_of_combobox.addItem(new_etat)
            self.etat_colors[new_etat] = "#f0f0f0"
            QMessageBox.information(self, "Succès", f"État '{new_etat}' ajouté avec succès!")
    
    def maj_statut(self, message):
        self.status_label.setText(message)
        QApplication.processEvents()
    
    def show_tree_menu(self, pos):
        item = self.tree.itemAt(pos)
        if item:
            self.context_menu_item = item
            self.context_menu_column = self.tree.header().logicalIndexAt(pos.x())
            self.tree_menu.exec_(self.tree.viewport().mapToGlobal(pos))
    
    def copy_tree_cell(self):
        if hasattr(self, 'context_menu_item') and hasattr(self, 'context_menu_column'):
            clipboard = QApplication.clipboard()
            clipboard.setText(self.context_menu_item.text(self.context_menu_column))
    
    def closeEvent(self, event):
        try:
            if hasattr(self, 'df') and not self.df.empty:
                sauvegarde_fermeture = os.path.join(self.data_folder, "SGOF_sauvegarde_fermeture.xlsx")

                # Ensure required columns exist
                if "OBSERVATION" not in self.df.columns:
                    self.df["OBSERVATION"] = ""

                colonnes_finales = [
                    "ANCIENCODE", "DESIGNATION PRODUIT", "N° OF",
                    "REALISE", "QUANTITEOF", "Etat d'OF","Famille", "Le poids d'OF","OBSERVATION","HOLE"
                ]

                colonnes_existantes = [col for col in colonnes_finales if col in self.df.columns]

                self.df[colonnes_existantes].to_excel(sauvegarde_fermeture, index=False)
                print("✅ Sauvegarde fermeture OK")
            else:
                print("ℹ️ Aucun DataFrame chargé, aucune sauvegarde de fermeture.")
        except Exception as e:
            print("❌ Erreur sauvegarde fermeture :", e)
        finally:
            event.accept()

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    screen = QGuiApplication.primaryScreen().geometry()
    window = ApplicationTraitementSGOF()
    window.resize(int(screen.width() * 0.8), int(screen.height() * 0.8))
    window.move((screen.width() - window.width()) // 2, (screen.height() - window.height()) // 2)
    
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
