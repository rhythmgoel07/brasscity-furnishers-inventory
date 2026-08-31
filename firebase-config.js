// ============================================================================
// Your Firebase project's configuration values.
// These are safe to be public in a web app — they identify your project, they
// are not secret passwords. Access is protected by the Firestore security rules
// (firestore.rules), not by hiding these.
// ============================================================================
export const firebaseConfig = {
  apiKey: "AIzaSyAtTs0BQqrsZFzeJPNyT473T80iorq4hPg",
  authDomain: "brasscity-inventory.firebaseapp.com",
  projectId: "brasscity-inventory",
  storageBucket: "brasscity-inventory.firebasestorage.app",
  messagingSenderId: "1076240840526",
  appId: "1:1076240840526:web:f2718a29e048009c3a501f",
  measurementId: "G-2TFGVJ98SC",

  // Google OAuth client ID for the Sheets sync (README Part 5).
  // Leave as-is to keep sync disabled; the rest of the app works without it.
  googleClientId: "PASTE_YOUR_GOOGLE_CLIENT_ID",
};
