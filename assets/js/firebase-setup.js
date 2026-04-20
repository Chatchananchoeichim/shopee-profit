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
    // 1. แปลงอีเมลเพื่อใช้ไปค้นหาใน Database (เปลี่ยน . เป็น _)
    const formattedEmail = user.email.replace(/\./g, '_');

    // 2. ไปเช็คที่ Node 'admin_users' ว่ามีชื่ออีเมลนี้และเป็น true หรือไม่
    firebase.database().ref('admin_users/' + formattedEmail).once('value')
      .then((snapshot) => {
        const isAdmin = snapshot.val() === true;

        if (isAdmin) {
          // ==========================================
          // โค้ดเดิมที่จะทำงานเมื่อเป็น Admin (อนุญาตให้เข้าแอป)
          // ==========================================
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
          // ==========================================
          // กรณีไม่ใช่ Admin (ไม่มีชื่อ หรือค่าไม่ใช่ true)
          // ==========================================
          showErrorMessage('ไม่อนุญาตให้เข้าใช้', 'อีเมลนี้ไม่มีสิทธิ์เข้าถึงระบบ กรุณาติดต่อผู้ดูแลระบบ');
          
          // บังคับ Logout ออกไป
          firebase.auth().signOut(); 
        }
      })
      .catch((error) => {
        console.error("Error checking admin permission:", error);
        showErrorMessage('ข้อผิดพลาด', 'ไม่สามารถตรวจสอบสิทธิ์ได้: ' + error.message);
        firebase.auth().signOut();
      });

  } else {
    // ==========================================
    // โค้ดเดิมที่จะทำงานตอนยังไม่ได้ Login หรือเพิ่ง Logout
    // ==========================================
    state.currentUser = null;
    if (typeof inactivityTimer !== 'undefined') clearTimeout(inactivityTimer);
    document.getElementById('auth-container').style.display = 'flex';
    document.getElementById('app-container').style.display = 'none';
    if(typeof dbRef !== 'undefined' && dbRef) dbRef.off(); // stop listening
  }
});

function toggleFirebaseSetup() {}
function connectFirebase(isAuto = false) { }
function disconnectFirebase() { }
