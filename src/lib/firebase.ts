import { initializeApp } from 'firebase/app';
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut } from 'firebase/auth';
import { 
  getFirestore, 
  collection, 
  doc, 
  setDoc, 
  getDoc, 
  getDocs, 
  updateDoc, 
  deleteDoc, 
  query, 
  where, 
  orderBy,
  serverTimestamp,
  getDocFromServer
} from 'firebase/firestore';
import firebaseConfig from '../../firebase-applet-config.json';

// Initialize Firebase services
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app, firebaseConfig.firestoreDatabaseId);
export const auth = getAuth();
export const googleProvider = new GoogleAuthProvider();

// Standard TS Interfaces
export interface FirebaseSeccion {
  id: string;
  userId: string;
  grado: string;
  nivel: string;
  nombre: string;
  color?: string;
  createdAt?: any;
  updatedAt?: any;
}

export interface FirebaseEstudiante {
  id: string;
  name: string;
  role: 'comunicador' | 'adivinador';
  locked: boolean;
  createdAt?: any;
}

// Error Handling helpers
export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  }
}

export function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error Context: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// Validate connection
async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      console.error("Please check your Firebase configuration or connection.");
    }
  }
}
testConnection();

// --- Firestore CRUD Helpers ---

/**
 * Fetch sections for current user
 */
export async function getSecciones(userId: string): Promise<FirebaseSeccion[]> {
  const path = 'secciones';
  try {
    const q = query(
      collection(db, path),
      where('userId', '==', userId)
    );
    const snap = await getDocs(q);
    const list: FirebaseSeccion[] = [];
    snap.forEach((d) => {
      list.push({ id: d.id, ...d.data() } as FirebaseSeccion);
    });
    return list;
  } catch (error) {
    handleFirestoreError(error, OperationType.LIST, path);
    return [];
  }
}

/**
 * Create or edit section details
 */
export async function saveSeccion(seccion: Partial<FirebaseSeccion> & { id: string; userId: string }, isEdit?: boolean) {
  const path = `secciones/${seccion.id}`;
  try {
    const docRef = doc(db, 'secciones', seccion.id);
    const now = serverTimestamp();
    
    if (isEdit) {
      await updateDoc(docRef, {
        grado: seccion.grado,
        nivel: seccion.nivel,
        nombre: seccion.nombre,
        color: seccion.color || 'indigo',
        updatedAt: now
      });
    } else {
      await setDoc(docRef, {
        userId: seccion.userId,
        grado: seccion.grado,
        nivel: seccion.nivel,
        nombre: seccion.nombre,
        color: seccion.color || 'indigo',
        createdAt: now,
        updatedAt: now
      });
    }
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

/**
 * Delete a section and all its students
 */
export async function deleteSeccionAndStudents(seccionId: string) {
  const sectionPath = `secciones/${seccionId}`;
  try {
    // 1. Delete students subcollection
    const studentsPath = `secciones/${seccionId}/estudiantes`;
    const sSnap = await getDocs(collection(db, 'secciones', seccionId, 'estudiantes'));
    const deletePromises: Promise<void>[] = [];
    sSnap.forEach((sd) => {
      deletePromises.push(deleteDoc(doc(db, 'secciones', seccionId, 'estudiantes', sd.id)));
    });
    await Promise.all(deletePromises);

    // 2. Delete parent section
    await deleteDoc(doc(db, 'secciones', seccionId));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, sectionPath);
  }
}

/**
 * Get students for a specific section
 */
export async function getEstudiantes(seccionId: string): Promise<FirebaseEstudiante[]> {
  const path = `secciones/${seccionId}/estudiantes`;
  try {
    const snap = await getDocs(collection(db, 'secciones', seccionId, 'estudiantes'));
    const list: FirebaseEstudiante[] = [];
    snap.forEach((d) => {
      const data = d.data();
      list.push({
        id: d.id,
        name: data.name,
        role: data.role || 'comunicador',
        locked: !!data.locked,
        createdAt: data.createdAt
      });
    });
    return list;
  } catch (error) {
    handleFirestoreError(error, OperationType.LIST, path);
    return [];
  }
}

/**
 * Save or insert student under a section
 */
export async function saveEstudiante(seccionId: string, estudiante: FirebaseEstudiante) {
  const path = `secciones/${seccionId}/estudiantes/${estudiante.id}`;
  try {
    const docRef = doc(db, 'secciones', seccionId, 'estudiantes', estudiante.id);
    await setDoc(docRef, {
      id: estudiante.id,
      name: estudiante.name,
      role: estudiante.role || 'comunicador',
      locked: estudiante.locked || false,
      createdAt: estudiante.createdAt || serverTimestamp()
    });
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

/**
 * Delete a student from a section
 */
export async function deleteEstudiante(seccionId: string, estudianteId: string) {
  const path = `secciones/${seccionId}/estudiantes/${estudianteId}`;
  try {
    await deleteDoc(doc(db, 'secciones', seccionId, 'estudiantes', estudianteId));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
}
