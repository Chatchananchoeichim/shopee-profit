// Initialize Firebase automatically with the hardcoded config
const firebaseConfig = {
  apiKey: "AIzaSyAmtmEQEFEQx4zf2Q1bbeTnhM3iPwPdnh4",
  authDomain: "torque-fd288.firebaseapp.com",
  databaseURL: "https://torque-fd288-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "torque-fd288",
  storageBucket: "torque-fd288.firebasestorage.app",
  messagingSenderId: "1040884189232",
  appId: "1:1040884189232:web:4a729a7f01a89df4da45f5"
};

if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
}
let dbRef = firebase.database().ref('shopee_cost_data');

// Authenticate via Email/Password
function login() {
  const email = document.getElementById('auth-email').value.trim();
  const pass = document.getElementById('auth-pass').value.trim();
  if(!email || !pass) {
    showWarningMessage('ข้อมูลไม่ครบ', 'กรุณากรอกอีเมลและรหัสผ่าน'); 
    return;
  }
  
  const btn = document.getElementById('login-btn');
  btn.disabled = true;
  btn.innerText = "กำลังเข้าสู่ระบบ...";

  // Set persistence to SESSION (cleared when browser closes)
  firebase.auth().setPersistence(firebase.auth.Auth.Persistence.SESSION)
    .then(() => {
      return firebase.auth().signInWithEmailAndPassword(email, pass);
    })
    .catch((error) => {
      showErrorMessage('เข้าสู่ระบบไม่สำเร็จ', "สาเหตุ: " + error.message);
      btn.disabled = false;
      btn.innerText = "เข้าสู่ระบบ";
    });
}

function logout() {
  firebase.auth().signOut();
}

// Idle Timeout Logic
let inactivityTimer;
const INACTIVITY_LIMIT_MS = 30 * 60 * 1000; // 30 minutes

function resetInactivityTimer() {
  if (state.currentUser) {
    clearTimeout(inactivityTimer);
    inactivityTimer = setTimeout(() => {
      showWarningMessage('เซสชันหมดอายุ', 'ไม่มีการใช้งานเป็นเวลานาน (30 นาที) กรุณาเข้าสู่ระบบใหม่เพื่อความปลอดภัย');
      logout();
    }, INACTIVITY_LIMIT_MS);
  }
}

// Track Auth State changes
firebase.auth().onAuthStateChanged((user) => {
  if (user) {
    state.currentUser = user;
    document.getElementById('auth-container').style.display = 'none';
    document.getElementById('app-container').style.display = 'flex';
    if(document.getElementById('user-email-sidebar')) {
      document.getElementById('user-email-sidebar').innerText = user.email;
    }
    if(document.getElementById('sync-status-text')) {
      document.getElementById('sync-status-text').innerText = 'ออนไลน์ซิงก์เรียบร้อยแล้ว (Cloud)';
    }
    if(document.getElementById('sync-status-dot')) {
      document.getElementById('sync-status-dot').style.background = 'var(--green)';
    }
    initCostMap(); 
    
    // Start idle timer and attach listeners
    resetInactivityTimer();
    ['mousemove', 'mousedown', 'keydown', 'touchstart'].forEach(evt => {
      document.addEventListener(evt, resetInactivityTimer);
    });
  } else {
    state.currentUser = null;
    clearTimeout(inactivityTimer);
    document.getElementById('auth-container').style.display = 'flex';
    document.getElementById('app-container').style.display = 'none';
    if(dbRef) dbRef.off(); // stop listening
  }
});

function toggleFirebaseSetup() {}
function connectFirebase(isAuto = false) { }
function disconnectFirebase() { }
