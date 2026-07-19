# Production Dashboard V3.3

גרסה זו מתקנת סנכרון בין מחשבים:
- באתר החי הנתונים נטענים קודם ורק מ-Netlify Blobs.
- בעת כניסת מנהל מתבצע איחוד בין הנתונים המקומיים במחשב הראשי לבין הנתונים שכבר בענן.
- האיחוד אינו מוחק רשומות קיימות ומונע כפילויות לפי מזהה/מאפייני הרשומה.
- נתוני איכות מקומיים מתאחדים גם הם עם הענן.


## V3.4
Fixed Netlify Blobs shared storage by explicitly using NETLIFY_SITE_ID and NETLIFY_TOKEN.
