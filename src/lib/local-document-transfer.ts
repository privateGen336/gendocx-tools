// Same-origin, local-only handoff between Compare and an Editor tab.
const DATABASE = "gendocx-document-handoff";
const STORE = "documents";
const TTL = 60 * 60 * 1000;

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DATABASE, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(STORE).createIndex("createdAt", "createdAt");
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function storeEditorDocument(file: File): Promise<string> {
  const database = await openDatabase();
  const id = crypto.randomUUID();
  try {
    await new Promise<void>((resolve, reject) => {
      const transaction = database.transaction(STORE, "readwrite");
      const store = transaction.objectStore(STORE);
      const expired = store.index("createdAt").openCursor(IDBKeyRange.upperBound(Date.now() - TTL));
      expired.onsuccess = () => {
        const cursor = expired.result;
        if (cursor) { cursor.delete(); cursor.continue(); }
      };
      store.put({ file, name: file.name, createdAt: Date.now() }, id);
      transaction.oncomplete = () => resolve();
      transaction.onabort = () => reject(transaction.error);
      transaction.onerror = () => reject(transaction.error);
    });
    return id;
  } finally { database.close(); }
}

export async function readEditorDocument(id: string): Promise<File> {
  const database = await openDatabase();
  try {
    const record = await new Promise<{ file: Blob; name: string; createdAt: number } | undefined>((resolve, reject) => {
      const request = database.transaction(STORE).objectStore(STORE).get(id);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    if (!record || Date.now() - record.createdAt > TTL) throw new Error("Document handoff expired");
    return new File([record.file], record.name, { type: record.file.type });
  } finally { database.close(); }
}
