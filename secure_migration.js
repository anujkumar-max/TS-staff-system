const admin = require('firebase-admin');
const fs = require('fs');

// IMPORTANT: Download your Firebase serviceAccountKey.json from the Firebase Console 
// (Project Settings -> Service Accounts -> Generate new private key)
// Place it in the same directory as this script.
const SERVICE_ACCOUNT_PATH = './serviceAccountKey.json';

if (!fs.existsSync(SERVICE_ACCOUNT_PATH)) {
  console.error(`ERROR: Could not find ${SERVICE_ACCOUNT_PATH}. Please download it from Firebase Console.`);
  process.exit(1);
}

const serviceAccount = require(SERVICE_ACCOUNT_PATH);

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();
const auth = admin.auth();

async function migrateUsers() {
  console.log("🚀 Starting Secure Firebase Auth Migration...");

  try {
    const staffSnapshot = await db.collection('staff').get();
    
    if (staffSnapshot.empty) {
      console.log("No staff records found.");
      return;
    }

    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;

    for (const doc of staffSnapshot.docs) {
      const data = doc.data();
      const phone = String(data.Phone || data.phone || '').trim();
      const rawPassword = String(data.password || '').trim();
      const existingRole = String(data.role || 'staff').trim();

      if (!phone || !rawPassword) {
        console.log(`⚠️ Skipping document ${doc.id} - Missing phone or password.`);
        skipCount++;
        continue;
      }

      // We append @appolice.local to create a unique fake email for Firebase Email/Password auth
      const email = `${phone}@appolice.local`;

      // Determine new PRO-MAP role
      let newRole = 'STAFF_PENDING_REVIEW';
      if (existingRole === 'admin') {
        newRole = 'ADMIN';
      }

      try {
        // 1. Create or get user in Firebase Auth
        let userRecord;
        try {
          userRecord = await auth.getUserByEmail(email);
          console.log(`ℹ️ User ${email} already exists in Firebase Auth.`);
        } catch (error) {
          if (error.code === 'auth/user-not-found') {
            userRecord = await auth.createUser({
              email: email,
              emailVerified: true,
              password: rawPassword, // Securely hashes it in Firebase Auth
              displayName: data.Full_Name || 'Staff',
              disabled: false,
            });
            console.log(`✅ Created Firebase Auth for ${email}`);
          } else {
            throw error;
          }
        }

        // 2. Create the promap_users document mapped to the Firebase Auth UID
        const promapUserRef = db.collection('promap_users').doc(userRecord.uid);
        const promapUserSnap = await promapUserRef.get();
        
        if (!promapUserSnap.exists) {
          await promapUserRef.set({
            authUid: userRecord.uid,
            userId: `TS-USR-${Date.now()}-${Math.floor(Math.random()*1000)}`,
            personnelId: '', // To be linked manually later if needed
            role: newRole,
            projectScopes: [], // Empty array by default
            active: true,
            createdAt: admin.firestore.FieldValue.serverTimestamp(),
            updatedAt: admin.firestore.FieldValue.serverTimestamp(),
            lastLoginAt: null
          });
          console.log(`✅ Created promap_users record for ${email}`);
        } else {
          console.log(`ℹ️ promap_users record already exists for ${email}`);
        }
        
        successCount++;

      } catch (err) {
        console.error(`❌ Failed to migrate ${phone}:`, err.message);
        errorCount++;
      }
    }

    console.log("\n=================================");
    console.log("🏁 MIGRATION COMPLETE");
    console.log(`✅ Successfully Migrated: ${successCount}`);
    console.log(`⚠️ Skipped: ${skipCount}`);
    console.log(`❌ Errors: ${errorCount}`);
    console.log("=================================");
    
    // IMPORTANT: Note that we DO NOT DELETE the legacy passwords here.
    // They are kept safely in the 'staff' collection until the transition is 100% verified.

  } catch (error) {
    console.error("Critical Error during migration:", error);
  }
}

migrateUsers();
