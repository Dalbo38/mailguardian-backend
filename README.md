# Mail Guardian - Backend

Backend Node.js pour la connexion IMAP à Outlook/Hotmail.

## Déploiement sur Render (gratuit)

1. Crée un compte sur https://render.com
2. "New Web Service" → connecte ton repo GitHub
3. Paramètres :
   - **Build Command** : `npm install`
   - **Start Command** : `npm start`
   - **Plan** : Free
4. Clique "Create Web Service"
5. Note l'URL fournie par Render (ex: https://mailguardian-backend.onrender.com)

## API

### POST /api/emails
Récupère les emails via IMAP.

Body JSON :
```json
{
  "email": "albo38@hotmail.fr",
  "password": "mot-de-passe-application-16-caracteres",
  "limit": 20
}
```
