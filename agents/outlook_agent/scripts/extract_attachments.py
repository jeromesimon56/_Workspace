import argparse
import base64
import email
import imaplib
import os
import sys

DEFAULT_IMAP_SERVER = "imap-mail.outlook.com"
DEFAULT_IMAP_PORT = 993


def xoauth2_auth(mail: imaplib.IMAP4_SSL, username: str, access_token: str):
    auth_string = f"user={username}\x01auth=Bearer {access_token}\x01\x01"

    def auth_callback(challenge):
        return auth_string.encode()

    return mail.authenticate("XOAUTH2", auth_callback)


def save_attachments(msg, output_dir):
    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if not filename:
                continue

            filepath = os.path.join(output_dir, filename)
            counter = 1
            base, ext = os.path.splitext(filepath)
            while os.path.exists(filepath):
                filepath = f"{base}_{counter}{ext}"
                counter += 1

            with open(filepath, "wb") as f:
                f.write(part.get_payload(decode=True))
            print("Pièce jointe sauvegardée :", filepath)


def parse_args():
    parser = argparse.ArgumentParser(description="Extraire les pièces jointes via IMAP Outlook.")
    parser.add_argument("--email", required=True, help="Adresse email Outlook/Hotmail")
    parser.add_argument("--password", help="Mot de passe ou mot de passe d'application Outlook")
    parser.add_argument("--token", help="Jeton d'accès OAuth2 pour XOAUTH2")
    parser.add_argument("--server", default=DEFAULT_IMAP_SERVER, help="Serveur IMAP Outlook")
    parser.add_argument("--port", type=int, default=DEFAULT_IMAP_PORT, help="Port IMAP TLS")
    parser.add_argument("--folder", default="INBOX", help="Dossier IMAP à ouvrir")
    parser.add_argument("--output", default="attachments", help="Dossier où enregistrer les pièces jointes")
    return parser.parse_args()


def main():
    args = parse_args()

    if not args.password and not args.token:
        print("Erreur : fournissez soit --password soit --token.")
        print("Pour un compte Outlook/Hotmail, le mot de passe simple peut être bloqué. Dans ce cas, utilisez OAuth2.")
        sys.exit(1)

    if args.password and args.token:
        print("Erreur : fournissez uniquement un mode d'authentification à la fois : --password OU --token.")
        sys.exit(1)

    os.makedirs(args.output, exist_ok=True)

    mail = imaplib.IMAP4_SSL(args.server, args.port)
    try:
        if args.token:
            print("Authentification XOAUTH2...")
            xoauth2_auth(mail, args.email, args.token)
        else:
            print("Authentification LOGIN...")
            mail.login(args.email, args.password)
    except imaplib.IMAP4.error as exc:
        print(f"IMAP auth failed to {args.server}:{args.port}.")
        print("Vérifie : 1) l'adresse, 2) le mot de passe/app-password ou 3) le jeton OAuth2.")
        print("Pour Outlook.com/Hotmail, Microsoft peut bloquer la basic auth. Si tu as MFA, utilise un mot de passe d'application ou OAuth2.")
        print(f"Détail : {exc}")
        sys.exit(1)

    typ, data = mail.select(args.folder)
    if typ != "OK":
        print(f"Impossible de sélectionner le dossier {args.folder}: {data}")
        mail.logout()
        sys.exit(1)

    typ, messages = mail.search(None, "ALL")
    if typ != "OK":
        print("Recherche IMAP échouée.")
        mail.logout()
        sys.exit(1)

    mail_ids = messages[0].split()
    for num in mail_ids:
        typ, data = mail.fetch(num, "(RFC822)")
        if typ != "OK":
            continue

        msg = email.message_from_bytes(data[0][1])
        save_attachments(msg, args.output)

    mail.logout()


if __name__ == "__main__":
    main()
