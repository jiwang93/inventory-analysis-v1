import { initializeApp, getApps } from 'firebase/app';
import { 
  initializeFirestore, 
  getFirestore, 
  collection, 
  addDoc, 
  getDocs, 
  query, 
  orderBy, 
  limit, 
  doc, 
  setDoc, 
  getDoc, 
  deleteDoc 
} from 'firebase/firestore';
import { getAuth, signInAnonymously } from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';

// Initialize Firebase App
const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApps()[0];
const firestoreDbId = (firebaseConfig as any).firestoreDatabaseId;

// Initialize Firestore with long-polling fallback for iframe/sandbox environments
let dbInstance;
try {
  dbInstance = initializeFirestore(app, {
    experimentalForceLongPolling: true,
  }, firestoreDbId || undefined);
} catch (e) {
  dbInstance = firestoreDbId ? getFirestore(app, firestoreDbId) : getFirestore(app);
}

export const db = dbInstance;
export const auth = getAuth(app);

// Helper to ensure authenticated session for Firestore reads/writes
export const ensureAuth = async () => {
  try {
    if (!auth.currentUser) {
      await signInAnonymously(auth);
    }
    return auth.currentUser;
  } catch (err) {
    console.warn('Anonymous auth initialization warning:', err);
    return null;
  }
};

// Eagerly trigger sign-in when loaded in browser
if (typeof window !== 'undefined') {
  ensureAuth();
}

export { collection, addDoc, getDocs, query, orderBy, limit, doc, setDoc, getDoc, deleteDoc };

