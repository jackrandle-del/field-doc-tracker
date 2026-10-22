import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager } from "firebase/firestore";

// Values come from Vercel env vars (Preview + Production), not hardcoded — see
// project_earthcraft memory / CLAUDE.md for the exact variable names to set there.
const firebaseConfig = {
  apiKey: process.env.REACT_APP_FIREBASE_API_KEY,
  authDomain: process.env.REACT_APP_FIREBASE_AUTH_DOMAIN,
  projectId: process.env.REACT_APP_FIREBASE_PROJECT_ID,
  storageBucket: process.env.REACT_APP_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: process.env.REACT_APP_FIREBASE_MESSAGING_SENDER_ID,
  appId: process.env.REACT_APP_FIREBASE_APP_ID,
};

const app = initializeApp(firebaseConfig);

export const auth = getAuth(app);

// Persistent local cache: a TA works on their phone with no signal in the field, then syncs once
// reconnected. No simultaneous multi-device editing is required (confirmed with the user), so
// Firestore's default last-write-wins merge on reconnect needs no custom conflict handling.
// Multi-tab on purpose: with the single-tab manager, a second open copy of the app (e.g. a link
// opened from Slack while the first is still open) silently fell back to a memory-only cache, so
// anything entered offline in that tab was lost when it closed. Multi-tab lets every open copy
// share the one saved cache and its queue of waiting writes.
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
});
