import getpass
import win32security
import win32con
import os

def check_windows_credentials(username, password):
    try:
        if "\\" in username:
            domain, user = username.split("\\", 1)
        else:
            domain = os.environ.get("USERDOMAIN", "")
            user = username

        print(f"Trying login for: {domain}\\{user}")
        print(f"Password length: {len(password)}")

        token = win32security.LogonUser(
            user,
            domain,
            password,
            win32con.LOGON32_LOGON_INTERACTIVE,  # Plus fiable
            win32con.LOGON32_PROVIDER_DEFAULT
        )
        token.Close()
        return True
    except Exception as e:
        print("[AUTH ERROR]", e)
        return False

if __name__ == "__main__":
    print("=== Test de connexion Windows ===")

    # Nom d'utilisateur automatique
    user = getpass.getuser()
    print(f"Nom d'utilisateur détecté : {user}")

    # Ou possibilité de le changer :
    change = input("Souhaitez-vous utiliser un autre utilisateur ? (y/n): ").strip().lower()
    if change == "y":
        user = input("Entrer le nom d'utilisateur (ou domaine\\utilisateur) : ")

    # Demande mot de passe caché
    password = getpass.getpass("Entrer le mot de passe : ")

    # Test
    if check_windows_credentials(user, password):
        print("✅ Connexion réussie !")
    else:
        print("❌ Nom d'utilisateur ou mot de passe incorrect.")
