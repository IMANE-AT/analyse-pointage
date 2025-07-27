# Fichier : db_logic.py

import sqlite3
import bcrypt

def init_db():
    """Initialise la base de données et crée la table des utilisateurs si elle n'existe pas."""
    with sqlite3.connect('data.db') as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                password TEXT NOT NULL
            )
        ''')

def add_user(username, password):
    """Ajoute un nouvel utilisateur avec un mot de passe haché."""
    if not username or not password:
        return False, "Le nom d'utilisateur et le mot de passe ne peuvent pas être vides."
        
    password_hash = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
    try:
        with sqlite3.connect('data.db') as conn:
            conn.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, password_hash))
        return True, "Compte créé avec succès !"
    except sqlite3.IntegrityError:
        return False, "Ce nom d'utilisateur existe déjà."

def check_user(username, password):
    """Vérifie si le mot de passe correspond à celui de l'utilisateur."""
    with sqlite3.connect('data.db') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT password FROM users WHERE username = ?", (username,))
        user_data = cursor.fetchone()

        if user_data:
            stored_password_hash = user_data[0]
            if bcrypt.checkpw(password.encode('utf-8'), stored_password_hash):
                return True
    return False

def check_if_users_exist():
    """Vérifie s'il y a au moins un utilisateur dans la base de données."""
    with sqlite3.connect('data.db') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users")
        count = cursor.fetchone()[0]
        return count > 0