import sys
import os
import pandas as pd
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QLineEdit, QMessageBox, QInputDialog, QFileDialog
)

class IHM(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ExcelMerger")
        self.resize(800, 400)
        self.css()
        
        # tableau
        self.table = QTableWidget(0, 4)
        self.table.setColumnWidth(0, 15)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(2, 300)
        self.table.setColumnWidth(3, 100)
        self.table.setHorizontalHeaderLabels(["", "", "Fichier", "Texte colonne 1"])
        self.table.horizontalHeader().setStretchLastSection(True)
        
        # boutons
        self.add_row_button = QPushButton("Ajouter une ligne")
        self.merge_button = QPushButton("Fusionner les fichiers")
        self.merge_button.setObjectName("btnMerge")
        self.save_config_button = QPushButton("Sauvegarder configuration")
        self.choose_config_button = QPushButton("Choisir configuration")
        
        self.add_row_button.clicked.connect(self.add_row)
        self.merge_button.clicked.connect(self.merge_files)
        self.save_config_button.clicked.connect(self.save_configuration)
        self.choose_config_button.clicked.connect(self.choose_configuration)
         
        # layout boutons
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_row_button)
        button_layout.addWidget(self.merge_button)
        button_layout.addWidget(self.save_config_button)
        button_layout.addWidget(self.choose_config_button)
        
        # layout principal
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.table)
        main_layout.addLayout(button_layout)
        
        # container principal
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
    
    
    def add_row(self):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        
        delete_button = QPushButton("×")
        delete_button.setObjectName("btnSupprLigne")
        delete_button.clicked.connect(lambda: self.remove_row(row_position))
        self.table.setCellWidget(row_position, 0, delete_button)
        
        file_button = QPushButton("Ajout fichier")
        file_button.clicked.connect(lambda: self.upload_file(row_position))
        self.table.setCellWidget(row_position, 1, file_button)
        
        file_path_edit = QLineEdit()
        self.table.setCellWidget(row_position, 2, file_path_edit)
        
        texte_premiere_colonne = QLineEdit()
        self.table.setCellWidget(row_position, 3, texte_premiere_colonne)
    
    
    def remove_row(self, row):
        self.table.removeRow(row)

    
    def upload_file(self, row):
        file_path, _ = QFileDialog.getOpenFileName(self, "Sélectionnez un fichier", "", "Fichiers Excel (*.xlsx)")
        if file_path:
            file_path_edit = self.table.cellWidget(row, 2)
            if isinstance(file_path_edit, QLineEdit):
                file_path_edit.setText(file_path)
            
            texte_premiere_colonne = self.table.cellWidget(row, 3)
            if isinstance(texte_premiere_colonne, QLineEdit):
                texte = os.path.splitext(os.path.basename(file_path))[0]
                texte_premiere_colonne.setText(texte)


    def get_all_file_paths_and_texts(self):
        files_and_texts = []
        row_count = self.table.rowCount()
        for row in range(row_count):
            file_path_edit = self.table.cellWidget(row, 2)
            texte_premiere_colonne = self.table.cellWidget(row, 3)
            if isinstance(file_path_edit, QLineEdit) and isinstance(texte_premiere_colonne, QLineEdit):
                file_path = file_path_edit.text()
                texte = texte_premiere_colonne.text()
                if file_path:
                    files_and_texts.append((file_path, texte))
        return files_and_texts


    def merge_files(self):
        files_and_texts = self.get_all_file_paths_and_texts()
        if not files_and_texts:
            QMessageBox.critical(self, "Erreur", "Aucun fichier à fusionner.")
            return
        
        combined_df = pd.DataFrame()
        columns_set = None
        
        for file, texte in files_and_texts:
            try:
                df = pd.read_excel(file)
        
                if columns_set is None:
                    columns_set = set(df.columns)
                elif set(df.columns) != columns_set:
                    QMessageBox.critical(self, "Erreur", f"Les colonnes du fichier {file} ne correspondent pas.")
                    return
                
                df.insert(0, '', texte)
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            
            except Exception as e:
                QMessageBox.critical(self, "Erreur", f"Erreur lors de la lecture du fichier {file}: {e}")
                return
        
        output_file, _ = QFileDialog.getSaveFileName(self, "Enregistrer le fichier fusionné", "", "Excel Files (*.xlsx)")
        if not output_file:
            return
        
        try:
            combined_df.to_excel(output_file, index=False)
            QMessageBox.information(self, "Succès", f"Fichiers fusionnés sauvegardés dans : {output_file}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la sauvegarde du fichier : {e}")


    def save_configuration(self):
        files_and_texts = self.get_all_file_paths_and_texts()
        if not files_and_texts:
            QMessageBox.critical(self, "Erreur", "Aucune configuration à sauvegarder.")
            return

        config_name, ok = QInputDialog.getText(self, "Nom de la configuration", "Entrez un nom pour la configuration :")
        if not ok or not config_name.strip():
            QMessageBox.critical(self, "Erreur", "Le nom de la configuration est requis.")
            return

        config_name = config_name.strip()

        try:
            existing_configs = []
            if os.path.exists("configs.txt"):
                with open("configs.txt", "r") as f:
                    existing_configs = [json.loads(line.strip()) for line in f.readlines()]
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la lecture des configurations existantes : {e}")
            return

        existing_names = [config["name"] for config in existing_configs]
        if config_name in existing_names:
            QMessageBox.critical(self, "Erreur", f"Une configuration avec le nom '{config_name}' existe déjà.")
            return

        config_line = {"name": config_name, "data": files_and_texts}
        try:
            with open("configs.txt", "a") as f:
                f.write(json.dumps(config_line) + "\n")
            QMessageBox.information(self, "Succès", f"Configuration '{config_name}' sauvegardée.")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la sauvegarde de la configuration : {e}")


    def choose_configuration(self):
        try:
            with open("configs.txt", "r") as f:
                lines = f.readlines()
            configs = [json.loads(line.strip()) for line in lines]
        except FileNotFoundError:
            QMessageBox.critical(self, "Erreur", "Aucune configuration n'a été trouvée.")
            return
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la lecture des configurations : {e}")
            return
        
        config_names = [config["name"] for config in configs]
        selected_config, ok = QInputDialog.getItem(self, "Choisir une configuration", "Sélectionnez une configuration :", config_names, editable=False)
        if not ok or not selected_config:
            return
        
        for config in configs:
            if config["name"] == selected_config:
                self.table.setRowCount(0)  # Réinitialise le tableau
                for file_path, texte in config["data"]:
                    self.add_row()
                    row_position = self.table.rowCount() - 1
                    file_path_edit = self.table.cellWidget(row_position, 2)
                    texte_premiere_colonne = self.table.cellWidget(row_position, 3)
                    if isinstance(file_path_edit, QLineEdit) and isinstance(texte_premiere_colonne, QLineEdit):
                        file_path_edit.setText(file_path)
                        texte_premiere_colonne.setText(texte)
                return


    def css(self):
        try:
            with open("styles.css", "r") as file:
                self.setStyleSheet(file.read())
        except FileNotFoundError:
            print("Le fichier styles.css est introuvable.")
            
            
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = IHM()
    window.show()
    sys.exit(app.exec())