// 1) Go to https://console.firebase.google.com
// 2) Create a project (free) -> Build -> Firestore Database -> Create database
//    (choose "Start in test mode" for now, or use the security rules below)
// 3) Project settings (gear icon) -> General -> "Your apps" -> Web app (</>) -> register app
// 4) Copy the firebaseConfig object Firebase gives you and paste its values below.

export const firebaseConfig = {
    apiKey: "AIzaSyA8iRE4U6EN7yevjWC8saGgo09E-zDHo2A",
    authDomain: "basic-442e4.firebaseapp.com",
    projectId: "basic-442e4",
    storageBucket: "basic-442e4.firebasestorage.app",
    messagingSenderId: "101682937860",
    appId: "1:101682937860:web:2a2d135bf5008d55057551",
    measurementId: "G-6BXHXRL9Q4"
};

// Firestore security rules (Firestore console -> Rules tab -> paste this -> Publish).
// Requires a signed-in Firebase Auth session (storage.js signs every browser in
// anonymously and automatically — no extra login screen). This blocks anyone who
// isn't running the actual app from reading/writing the database directly, even
// though the app's own username/password screen is a separate, stronger layer on
// top of this for who gets into the UI itself.
//
// Before this works, enable it once: Firebase console -> Authentication ->
// Sign-in method -> enable "Anonymous".
/*
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /storage/{docId} {
      allow read, write: if request.auth != null;
    }
  }
}
*/
