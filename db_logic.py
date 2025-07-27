# Fichier : db_logic.py
import sqlite3
import bcrypt
import secrets
import datetime

def init_db():
    """Initialise la DB et ajoute les colonnes pour le reset de mot de passe."""
    with sqlite3.connect('data.db') as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                password TEXT NOT NULL,
                reset_token TEXT,
                token_expiry TIMESTAMP
            )
        ''')

def add_user(username, password):
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
    with sqlite3.connect('data.db') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT password FROM users WHERE username = ?", (username,))
        user_data = cursor.fetchone()
        if user_data:
            return bcrypt.checkpw(password.encode('utf-8'), user_data[0])
    return False

def update_password(username, old_password, new_password):
    if not new_password: return False, "Le nouveau mot de passe ne peut pas être vide."
    if check_user(username, old_password):
        new_password_hash = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
        with sqlite3.connect('data.db') as conn:
            conn.execute("UPDATE users SET password = ? WHERE username = ?", (new_password_hash, username))
        return True, "Mot de passe mis à jour avec succès !"
    return False, "L'ancien mot de passe est incorrect."

def set_reset_token(username):
    """Génère et stocke un token de reset pour un utilisateur."""
    with sqlite3.connect('data.db') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT username FROM users WHERE username = ?", (username,))
        if cursor.fetchone():
            token = secrets.token_urlsafe(20)
            expiry = datetime.datetime.now() + datetime.timedelta(hours=1) # Le token expire dans 1 heure
            conn.execute("UPDATE users SET reset_token = ?, token_expiry = ? WHERE username = ?", (token, expiry, username))
            return token
    return None

def reset_password_with_token(token, new_password):
    """Réinitialise le mot de passe si le token est valide."""
    if not new_password: return False, "Le mot de passe ne peut être vide."
    with sqlite3.connect('data.db') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT username, token_expiry FROM users WHERE reset_token = ?", (token,))
        user_data = cursor.fetchone()
        if user_data:
            username, token_expiry_str = user_data
            token_expiry = datetime.datetime.fromisoformat(token_expiry_str)
            if token_expiry > datetime.datetime.now():
                new_password_hash = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
                conn.execute("UPDATE users SET password = ?, reset_token = NULL, token_expiry = NULL WHERE username = ?", (new_password_hash, username))
                return True, "Votre mot de passe a été réinitialisé avec succès."
    return False, "Le lien de réinitialisation est invalide ou a expiré."

def check_if_users_exist():
    with sqlite3.connect('data.db') as conn:
        return conn.execute("SELECT COUNT(*) FROM users").fetchone()[0] > 0