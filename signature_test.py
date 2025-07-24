import win32com.client as win32
import time

def get_outlook_signature():
    try:
        # Crée une instance Outlook
        outlook = win32.Dispatch("Outlook.Application")
        
        # Crée un email temporaire
        mail = outlook.CreateItem(0)  # 0 = Mail item (nouveau message)
        
        # Affiche l'email (permet d'ajouter automatiquement la signature)
        mail.Display()
        
        # Attendre que la signature soit insérée
        time.sleep(2)  # Attente de 2 secondes (donne le temps à Outlook d'ajouter la signature)
        
        # Récupère le corps de l'email (avec la signature)
        signature = mail.HTMLBody
        
        # Ferme le brouillon sans envoyer
        mail.Close(0)
        
        # Si la signature est vide ou incorrecte, renvoie un message par défaut
        if not signature or signature == "<BODY></BODY>":
            print("Signature vide ou non récupérée, envoie une signature par défaut.")
            signature = """
                <br><br>
                <p>Best regards,</p>
                <p>Your Name</p>
                <p>Your Position</p>
                <p>Your Company</p>
            """
        
        return signature
    
    except Exception as e:
        print(f"Erreur lors de la récupération de la signature : {e}")
        return ""

# Test de récupération de la signature
signature = get_outlook_signature()

# Affichage de la signature dans la console
print("Signature récupérée :")
print(signature)
