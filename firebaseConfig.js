// 1) Go to https://console.firebase.google.com
// 2) Create a project (free) -> Build -> Firestore Database -> Create database
//    (choose "Start in test mode" for now, or use the security rules below)
// 3) Project settings (gear icon) -> General -> "Your apps" -> Web app (</>) -> register app
// 4) Copy the firebaseConfig object Firebase gives you and paste its values below.

export const firebaseConfig = {
  apiKey: "AIzaSyBUoKVFXJRmLz69GrsBOTJs4S15-20WDYg",
  authDomain: "basic-3d4ac.firebaseapp.com",
  projectId: "basic-3d4ac",
  storageBucket: "basic-3d4ac.firebasestorage.app",
  messagingSenderId: "638212034626",
  appId: "1:638212034626:web:49a8bcd66bdf9fc00e5ddb"
};

// Recommended Firestore security rules (Firestore console -> Rules tab).
// This keeps the "storage" collection open to read/write, since the app already has
// its own username/password login screen guarding access to the UI. If you want
// stronger protection, look into Firebase Authentication + rules that check auth.uid.
/*
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /storage/{docId} {
      allow read, write: if true;
    }
  }
}
*/
