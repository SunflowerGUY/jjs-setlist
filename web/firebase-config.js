/* JJ's Setlist - online accounts (Firebase).

   To turn on online accounts, replace null below with your Firebase project's
   settings: Firebase console > Project settings > General > Your apps >
   (your web app) > SDK setup and configuration > Config. It looks like:

     window.JJS_FIREBASE_CONFIG = {
       apiKey: "...",
       authDomain: "your-project.firebaseapp.com",
       projectId: "your-project",
       storageBucket: "your-project.firebasestorage.app",
       messagingSenderId: "...",
       appId: "...",
     };

   These settings are not a password: they're meant to be in the web page.
   What protects everyone's data is firestore.rules (paste it into
   Firestore Database > Rules). Leave this as null to keep accounts off. */
window.JJS_FIREBASE_CONFIG = null;
