# גרסה 3.2 — אחסון נתונים משותף

1. העלה את כל קבצי הפרויקט ל-GitHub באותו מאגר.
2. ודא שהקובץ `netlify/functions/storage.js` עודכן.
3. ב-Netlify השאר את משתני הסביבה:
   - `ADMIN_CODE`
   - `ADMIN_TOKEN_SECRET`
4. המתן ל-Deploy מוצלח.
5. פתח את האתר במחשב שבו נמצאים הנתונים הישנים והיכנס כמנהל פעם אחת.
6. לאחר הכניסה המערכת תעביר את הנתונים הקיימים ל-Netlify Blobs.
7. פתח את האתר במחשב אחר ובצע Ctrl+F5. הנתונים אמורים להופיע.

אין צורך ב-SITE_ID או NETLIFY_TOKEN לצורך Netlify Blobs מתוך Netlify Functions.
