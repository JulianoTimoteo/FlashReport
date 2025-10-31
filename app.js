// --- Firebase Imports (Versão 10.7.1) ---
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, onAuthStateChanged, signOut, createUserWithEmailAndPassword, signInWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";
import { getFirestore, doc, setDoc, updateDoc, collection, query, onSnapshot, deleteDoc, where, addDoc, getDocs, setLogLevel, getDoc } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js";
import Sortable from "https://cdn.jsdelivr.net/npm/sortablejs@1.15.0/modular/sortable.esm.js";

// 1. Top-level Vars & State (Globals - acessíveis por todo o arquivo)
let currentFillReportId = null;
let confirmCallback = null;

// Permissões padrão
const allPermissions = { canViewDashboard: true, canCreateReports: true, canFillReports: true, canViewAnalysis: true, canManageUsers: true };
const analystPermissions = { canViewDashboard: true, canCreateReports: false, canFillReports: false, canViewAnalysis: true, canManageUsers: false };
const userPermissions = { canViewDashboard: true, canCreateReports: false, canFillReports: true, canViewAnalysis: false, canManageUsers: false };

const state = {
    currentUser: null,
    currentReportComponents: [],
    savedReports: [],
    filledReports: [],
    draftReports: [],
    pendingUsers: [],
    notifications: { pendingUsers: 0, draftReports: 0 },
    users: [], // Usuários do Firestore
    currentPage: 'dashboard',
    darkMode: localStorage.getItem('darkMode') === 'true',
    editingComponentId: null,
    editingReportId: null,
    speechRecognition: null,
    speechTarget: null // Elemento de input atual para o ditado
};

// 2. Elementos DOM
const elements = {
    loginScreen: document.getElementById('login-screen'),
    appScreen: document.getElementById('main-app-view'),
    loginForm: document.getElementById('login-form'),
    loginError: document.getElementById('login-error'),
    logoutBtn: document.getElementById('logout-btn'),
    themeToggle: document.getElementById('theme-toggle'),
    menuToggle: document.getElementById('menu-toggle'),
    sidebar: document.getElementById('sidebar'),
    pageTitle: document.getElementById('page-title'),
    pageContent: document.getElementById('page-content'),
    toast: document.getElementById('toast'),
    toastMessage: document.getElementById('toast-message'),
    toastIcon: document.getElementById('toast')?.querySelector('i'),
    userModal: document.getElementById('user-modal'),
    userForm: document.getElementById('user-form'),
    userModalTitle: document.getElementById('user-modal-title'),
    cancelUserBtn: document.getElementById('cancel-user-btn'),
    addUserBtn: document.getElementById('add-user-btn'),
    reportBuilderArea: document.getElementById('report-builder-area'),
    reportPreviewArea: document.getElementById('report-preview-area'),
    previewTitle: document.getElementById('preview-title'),
    previewComponents: document.getElementById('report-preview-area'),
    reportTitle: document.getElementById('report-title'),
    saveReportBtn: document.getElementById('save-report-btn'),
    cancelEditBtn: document.getElementById('cancel-edit-btn'),
    editingReportId: document.getElementById('editing-report-id'),
    savedReportsList: document.getElementById('saved-reports-list'),
    filledReportsList: document.getElementById('filled-reports-list'),
    usersTableBody: document.getElementById('users-table-body'),
    componentModal: document.getElementById('component-modal'),
    componentForm: document.getElementById('component-form'),
    componentModalTitle: document.getElementById('component-modal-title'),
    componentId: document.getElementById('component-id'),
    componentLabel: document.getElementById('component-label'),
    componentOptionsContainer: document.getElementById('component-options-container'),
    componentOptions: document.getElementById('component-options'),
    componentColumnsContainer: document.getElementById('component-columns-container'),
    componentColumns: document.getElementById('component-columns'),
    componentRequired: document.getElementById('component-required'),
    componentSingleSelectionContainer: document.getElementById('component-single-selection-container'),
    componentSingleSelection: document.getElementById('component-single-selection'),
    cancelComponentBtn: document.getElementById('cancel-component-btn'),
    fillReportModal: document.getElementById('fill-report-modal'),
    fillReportModalTitle: document.getElementById('fill-report-modal-title'),
    fillReportForm: document.getElementById('fill-report-form'),
    cancelFillReportBtn: document.getElementById('cancel-fill-report-btn-modal'),
    confirmModalEl: document.getElementById('confirm-modal'),
    confirmTitleEl: document.getElementById('confirm-modal-title'),
    confirmMessageEl: document.getElementById('confirm-modal-message'),
    confirmBtn: document.getElementById('confirm-modal-confirm'),
    confirmCancelBtn: document.getElementById('confirm-modal-cancel'), 
    userInfoSidebar: document.getElementById('user-info-sidebar'),
    userAvatarSidebar: document.getElementById('user-avatar-sidebar'),
    userNameSidebar: document.getElementById('user-name-sidebar'),
    userEmailSidebar: document.getElementById('user-email-sidebar'),
    notificationBellButton: document.getElementById('notification-bell-button'),
    notificationDot: document.getElementById('notification-dot'),
    notificationDropdown: document.getElementById('notification-dropdown'),
    notificationList: document.getElementById('notification-list'),
    notificationEmptyState: document.getElementById('notification-empty-state'),
    requestAccessLink: document.getElementById('request-access-link'),
    pendingUsersContainer: document.getElementById('pending-users-container'),
    pendingUsersTableBody: document.getElementById('pending-users-table-body'),
    saveDraftBtn: document.getElementById('save-draft-btn-modal'),
    draftsListContainer: document.getElementById('drafts-list-container'),
    draftsList: document.getElementById('drafts-list'),
    viewReportModal: document.getElementById('view-report-modal'),
    viewReportModalTitle: document.getElementById('view-report-modal-title'),
    viewReportFilledBy: document.getElementById('view-report-filled-by'),
    viewReportFilledAt: document.getElementById('view-report-filled-at'),
    viewReportContent: document.getElementById('view-report-content'),
    closeViewReportBtn: document.getElementById('close-view-report-btn'),
    userPasswordInput: document.getElementById('user-password'), 
    togglePasswordVisibility: null,
    loginEmailInput: document.getElementById('username'),
    loginPasswordInput: document.getElementById('password'),
    toggleLoginPasswordVisibility: null,
};

// 3. --- Firebase Configuração (Chaves Reais) ---
const FIREBASE_CONFIG = {
    // CHAVES REAIS DO FIREBASE 
    apiKey: "AIzaSyDW0N4iBt7xfJvK7qC8VCiB56NukDhDN_A",
    authDomain: "superapp-relatorios.firebaseapp.com",
    projectId: "superapp-relatorios",
    storageBucket: "superapp-relatorios.firebasestorage.app",
    messagingSenderId: "882241804299",
    appId: "1:882241804299:web:21b4529ed930f06c6d8f0e"
};

// 4. --- Variáveis de Inicialização Firebase ---
let app, auth, db, googleProvider;
let isFirebaseReady = false;

// Paths de coleção
const APP_ID = FIREBASE_CONFIG.projectId; 
const ARTIFACT_ID = "superapp-relatorios-artifact"; 

const USERS_COLLECTION_PATH = `appId/${APP_ID}/public/data/users`;
const REPORT_TEMPLATES_COLLECTION_PATH = `artifacts/${ARTIFACT_ID}/public/data/report_templates`;
const FILLED_REPORTS_COLLECTION_PATH = `artifacts/${ARTIFACT_ID}/public/data/filled_reports`;
const DRAFTS_COLLECTION_PATH = `appId/${APP_ID}/public/data/drafts`;
const ACTIVITY_LOGS_COLLECTION_PATH = `appId/${APP_ID}/public/data/activity_logs`;


// ----------------------------------------------------------------------
// --- FUNÇÕES DE UI E UTILITY ------------------------------------------
// ----------------------------------------------------------------------

function showToast(message, type = 'success') {
    const toastEl = elements.toast; const iconEl = elements.toastIcon; const messageEl = elements.toastMessage;
    if (!toastEl || !iconEl || !messageEl) { console.warn("[Toast] Elements missing"); console.log("[Toast] Message:", message); return; }
    messageEl.textContent = message; let iconName = 'check-circle'; let iconClass = 'w-5 h-5 text-green-500';
    if (type === 'error') { iconName = 'x-circle'; iconClass = 'w-5 h-5 text-red-500'; }
    else if (type === 'warning') { iconName = 'alert-triangle'; iconClass = 'w-5 h-5 text-yellow-500'; }
    else if (type === 'info') { iconName = 'info'; iconClass = 'w-5 h-5 text-blue-500'; } 
    iconEl.setAttribute('data-lucide', iconName); iconEl.className = iconClass;
    try { lucide.createIcons({ nodes: [iconEl] }); } catch (e) { console.error("[Lucide Error] Toast Icon:", e); }
    toastEl.classList.add('show'); setTimeout(() => { toastEl.classList.remove('show'); }, 3000);
}
function getProfileBadgeClass(profile) {
    const classes = { admin: 'bg-red-100 text-red-800 dark:bg-red-900 dark:text-red-200', analyst: 'bg-purple-100 text-purple-800 dark:bg-purple-900 dark:text-purple-200', user: 'bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200'};
    return classes[profile] || 'bg-gray-100 dark:bg-gray-700';
}
function getProfileName(profile) {
    const names = { admin: 'Admin', analyst: 'Analista', user: 'Usuário' }; return names[profile] || profile;
}
function getPagePermissionKey(page) {
    const keys = { 'dashboard': 'canViewDashboard', 'create-report': 'canCreateReports', 'fill-report': 'canFillReports', 'analysis': 'canViewAnalysis', 'users': 'canManageUsers' };
    return keys[page] || 'canViewDashboard';
}
function updateNavActiveState(page) {
    document.querySelectorAll('#main-nav .nav-item').forEach(item => { item.classList.remove('active'); if (item.dataset.page === page) item.classList.add('active'); });
}
function togglePasswordVisibility(inputElement, buttonElement) {
    const isPassword = inputElement.type === 'password';
    inputElement.type = isPassword ? 'text' : 'password';
    buttonElement.querySelector('i')?.setAttribute('data-lucide', isPassword ? 'eye-off' : 'eye');
    try { lucide.createIcons({ nodes: [buttonElement.querySelector('i')] }); } catch (e) { console.error("[Lucide Error] Toggle icon:", e); }
}
function showConfirmModal(title, message, onConfirm, confirmButtonClass = 'btn-danger') {
    if (!elements.confirmModalEl) { console.error("[Modal] Confirm modal missing"); if (confirm(`${title}\n${message}`)) { if(typeof onConfirm === 'function') onConfirm(); } return; }
    elements.confirmTitleEl.textContent = title; elements.confirmMessageEl.textContent = message; confirmCallback = onConfirm;
    elements.confirmBtn.className = `btn ${confirmButtonClass}`; elements.confirmModalEl.classList.add('active');
    console.log(`[Modal] Confirm modal shown: ${title}`);
}
function closeConfirmModal() { if (elements.confirmModalEl) elements.confirmModalEl.classList.remove('active'); confirmCallback = null; }
function closeUserModal() { if(elements.userModal) elements.userModal.classList.remove('active'); }
function closeComponentModal() {
    if(elements.componentModal) elements.componentModal.classList.remove('active');
    console.log("[Modal] Component modal closed.");
}
function closeFillReportModal() {
    elements.fillReportModal.classList.remove('active'); 
    elements.fillReportForm.innerHTML = ''; 
    currentFillReportId = null; 
    elements.fillReportForm.dataset.draftId = '';
    
    // Remove a classe de scroll do modal
    elements.fillReportModal.querySelector('.modal')?.classList.remove('modal-fill-scroll');

    console.log("[Fill] Closed fill report modal.");
}
function closeViewReportModal() {
    if (elements.viewReportModal) elements.viewReportModal.classList.remove('active');
    console.log("[View] Closed view report modal.");
}
function toggleSidebar() { if(elements.sidebar) elements.sidebar.classList.toggle('open'); }
function toggleTheme() {
    state.darkMode = !state.darkMode; document.documentElement.classList.toggle('dark');
    if(elements.themeToggle) elements.themeToggle.classList.toggle('active'); localStorage.setItem('darkMode', state.darkMode);
    console.log(`[UI] Theme toggled. Dark mode: ${state.darkMode}`);
}
function updatePreviewTitle() {
    const title = elements.reportTitle?.value || 'Título do Relatório';
    if(elements.previewTitle) elements.previewTitle.textContent = title;
}
function updateNotificationBell() {
    const totalNotifications = state.notifications.pendingUsers + state.notifications.draftReports;
    if (totalNotifications > 0) { elements.notificationDot.classList.remove('hidden'); }
    else { elements.notificationDot.classList.add('hidden'); }
    renderNotificationDropdown();
}

// ----------------------------------------------------------------------
// --- FUNÇÕES DE NAVEGAÇÃO E AUTENTICAÇÃO (Integração com Firebase) ----
// ----------------------------------------------------------------------

function updateUserInfoSidebar(userData) {
    if (elements.userInfoSidebar && elements.userAvatarSidebar && elements.userNameSidebar && elements.userEmailSidebar) {
        if (userData) {
            elements.userNameSidebar.textContent = userData.name || 'Usuário'; elements.userEmailSidebar.textContent = userData.email || 'email@desconhecido.com';
            const avatarUrl = `https://ui-avatars.com/api/?name=${encodeURIComponent(userData.name || '?')}&background=random&color=fff`;
            elements.userAvatarSidebar.src = avatarUrl; elements.userInfoSidebar.classList.remove('hidden');
        } else {
            elements.userInfoSidebar.classList.add('hidden'); elements.userAvatarSidebar.src = "https://placehold.co/32x32/E0E7FF/334155?text=?";
        }
    } else { console.warn("[UI] User info sidebar elements not found."); }
}

function updateNavVisibility() {
     if (!state.currentUser) return;
     console.log(`[Nav] Updating visibility for user ${state.currentUser.name} (${state.currentUser.profile})`);
     document.querySelectorAll('#main-nav .nav-item').forEach(item => {
         const page = item.dataset.page;
         const permKey = getPagePermissionKey(page);
         if (state.currentUser.permissions[permKey]) {
             item.style.display = 'flex';
         } else {
             item.style.display = 'none';
         }
     });
}

function switchPage(page) {
    console.log(`[Nav] Switching page to: ${page}`);
    document.querySelectorAll('.page').forEach(p => p.classList.add('hidden'));
    const pageElement = document.getElementById(`page-${page}`);
    if (state.currentUser && !state.currentUser.permissions[getPagePermissionKey(page)]) {
        showToast("Acesso negado.", "error");
        console.warn(`[Nav] Access denied for user ${state.currentUser.name} to page ${page}`);
        document.getElementById(`page-dashboard`).classList.remove('hidden');
        state.currentPage = 'dashboard'; updateNavActiveState('dashboard'); return;
    }
    if (pageElement) {
        pageElement.classList.remove('hidden'); state.currentPage = page;
        const pageTitles = { 'dashboard': 'Dashboard', 'create-report': 'Criar Relatório', 'fill-report': 'Preencher Relatório', 'analysis': 'Análise', 'users': 'Usuários' };
        if(elements.pageTitle) elements.pageTitle.textContent = pageTitles[page] || 'Super App';
        updateNavActiveState(page);
        if (window.innerWidth < 1024 && elements.sidebar) elements.sidebar.classList.remove('open');
        if (page !== 'create-report' && state.editingReportId) {
            console.log("[Nav] Leaving create-report page while editing, cancelling edit.");
            cancelReportEdit(false);
        }
    } else { console.error(`[Nav] Page element not found: page-${page}. Switching to dashboard.`); switchPage('dashboard'); }
}

function handleLogin(e) {
    e.preventDefault();
    const email = elements.loginEmailInput.value;
    const password = elements.loginPasswordInput.value;
    elements.loginError.classList.add('hidden');

    const loginButton = elements.loginForm.querySelector('button[type="submit"]');

    if (email === 'admin' && password === 'admin') {
        const adminUser = { id: 'local_admin', name: 'Admin (Local)', email: 'admin@local.com', profile: 'admin', permissions: { ...allPermissions } };
        state.currentUser = adminUser;
        elements.loginScreen.classList.add('hidden'); elements.appScreen.classList.remove('hidden');
        updateUserInfoSidebar(state.currentUser); updateNotificationBell(); updateNavVisibility();
        
        if (!state.users.find(u => u.id === 'local_admin')) {
             state.users.unshift(adminUser);
        }
        
        showToast('Login como Admin (Local) realizado com sucesso! (Modo Offline)');
        renderUsersTable(); switchPage('dashboard');
        return;
    }

    if (isFirebaseReady) {
        loginButton.disabled = true;
        
        signInWithEmailAndPassword(auth, email, password)
            .catch(error => {
                console.error("[Auth] Email/Password Login Error:", error);
                elements.loginError.textContent = "Usuário/senha incorretos ou acesso não autorizado.";
                elements.loginError.classList.remove('hidden');
            })
            .finally(() => {
                loginButton.disabled = false;
                try { lucide.createIcons(); } catch(e) {}
            });
        return;
    }

    elements.loginError.textContent = "Firebase não inicializado ou credenciais incorretas.";
    elements.loginError.classList.remove('hidden');
}

function handleLogout() {
    showConfirmModal("Confirmar Saída", "Tem certeza que deseja sair?", () => {
        if (state.currentUser && state.currentUser.id !== 'local_admin' && isFirebaseReady) {
            firebaseHandleLogout(); 
        } else {
            elements.appScreen.classList.add('hidden'); elements.loginScreen.classList.remove('hidden');
            state.currentUser = null;
            elements.loginEmailInput.value = 'admin'; elements.loginPasswordInput.value = 'admin';
            updateUserInfoSidebar(null); showToast('Logout realizado com sucesso!');
            switchPage('dashboard');
        }
    }, 'btn-danger');
}

// ----------------------------------------------------------------------
// --- FUNÇÕES FIREBASE E OPERAÇÕES NO BANCO ----------------------------
// ----------------------------------------------------------------------

function initFirebase() {
    try {
        setLogLevel('Debug');
        app = initializeApp(FIREBASE_CONFIG);
        auth = getAuth(app);
        db = getFirestore(app);
        googleProvider = new GoogleAuthProvider();
        isFirebaseReady = true;

        state.users = state.users.filter(u => u.id === 'local_admin'); 
        state.savedReports = []; state.filledReports = []; state.draftReports = []; state.pendingUsers = [];
        state.notifications.pendingUsers = 0; state.notifications.draftReports = 0;
        
        onAuthStateChanged(auth, (user) => {
            if (user) {
                checkAuthorizationAndLogin(user);
            } else {
                console.log("[Auth] No user logged in (Firebase). Showing login screen.");
                state.currentUser = null;
                elements.loginScreen.classList.remove('hidden');
                elements.appScreen.classList.add('hidden');
                updateUserInfoSidebar(null);
                updateNavVisibility();
            }
        });
    } catch (e) {
        console.error("[Firebase] Initialization Error:", e);
        showToast("Erro fatal ao inicializar o Firebase. Verifique a chave API e o console.", "error");
    }
}

async function logActivity(action, details) {
    if (!state.currentUser || !isFirebaseReady || state.currentUser.id === 'local_admin') return;
    try {
        await addDoc(collection(db, ACTIVITY_LOGS_COLLECTION_PATH), {
            user: state.currentUser.email,
            name: state.currentUser.name,
            action: action,
            timestamp: new Date(),
            details: details
        });
    } catch (e) {
        console.error("[Firestore] Error logging activity:", e);
    }
}

async function checkAuthorizationAndLogin(user) {
    const userDocRef = doc(db, USERS_COLLECTION_PATH, user.uid);
    
    try {
        const userDoc = await getDoc(userDocRef);

        if (userDoc.exists()) {
            const userData = userDoc.data();
            
            if (userData.autorizado === true) {
                const profile = userData.profile || 'admin';
                const permissions = userData.permissions || (profile === 'admin' ? allPermissions : (profile === 'analyst' ? analystPermissions : userPermissions));

                state.currentUser = {
                    id: user.uid, 
                    name: userData.name || user.email.split('@')[0],
                    email: user.email,
                    profile,
                    permissions
                };

                elements.loginScreen.classList.add('hidden');
                elements.appScreen.classList.remove('hidden');
                updateUserInfoSidebar(state.currentUser);
                updateNotificationBell();
                updateNavVisibility();
                
                showToast(`Bem-vindo(a), ${state.currentUser.name.split(' ')[0]}!`);
                switchPage('dashboard');

                fetchReportTemplates();
                fetchDrafts();
                fetchFilledReports();
                if (state.currentUser.permissions.canManageUsers) {
                    fetchUsersRealtime(); 
                }

            } else {
                showToast("Acesso pendente. Aguarde aprovação do administrador.", "error");
                await signOut(auth);
            }
        } else {
            const newUserRecord = {
                name: user.displayName || user.email.split('@')[0],
                email: user.email,
                profile: 'user',
                autorizado: false, 
                createdAt: new Date().toISOString(),
                permissions: userPermissions
            };

            await setDoc(userDocRef, newUserRecord);
            showToast("Sua solicitação de acesso foi enviada. Aguarde aprovação.", "warning");
            await signOut(auth);
        }
    } catch (e) {
        console.error("[Firestore] Error fetching user profile:", e);
        showToast("Erro de acesso. Tente novamente ou contate o suporte.", "error");
        await signOut(auth).catch(e => console.error("SignOut failed:", e));
    }
}

async function signInWithGoogle() {
    if (!isFirebaseReady) { showToast("Firebase não inicializado.", "error"); return; }
    
    const googleBtn = document.getElementById('login-google-btn');
    if (googleBtn) googleBtn.disabled = true;

    try {
        await signInWithPopup(auth, googleProvider);
    } catch (error) {
        console.error("[Auth] Google Sign-In Error:", error);
        if (error.code !== 'auth/popup-closed-by-user') {
            showToast("Falha no login Google. Verifique se o domínio está autorizado.", "error");
        }
    } finally {
        if (googleBtn) googleBtn.disabled = false;
    }
}

async function firebaseHandleLogout() {
    if (!isFirebaseReady) throw new Error("Firebase não pronto.");
    try {
        await signOut(auth);
    } catch (error) {
        console.error("[Auth] Firebase Logout Error:", error);
        showToast("Erro ao sair.", "error");
    }
}

// --- Operações Firestore (CRUD) ---

async function saveReportTemplateToFirestore(id, data) {
    if (!isFirebaseReady) throw new Error("Firebase não pronto.");
    
    const docData = { ...data, userId: state.currentUser.id }; 

    const docRef = id ? doc(db, REPORT_TEMPLATES_COLLECTION_PATH, id) : doc(collection(db, REPORT_TEMPLATES_COLLECTION_PATH));
    await setDoc(docRef, docData, { merge: true });
    logActivity("save-template", { templateId: docRef.id, templateTitle: data.title });
    return docRef.id;
}
async function updateReportTemplateStatusInFirestore(id, status) {
    if (!isFirebaseReady) throw new Error("Firebase não pronto.");
    const docRef = doc(db, REPORT_TEMPLATES_COLLECTION_PATH, id);
    await updateDoc(docRef, { status: status });
    logActivity("update-template-status", { templateId: id, newStatus: status });
}
async function deleteReportTemplateFromFirestore(id) {
    if (!isFirebaseReady) throw new Error("Firebase não pronto.");
    const docRef = doc(db, REPORT_TEMPLATES_COLLECTION_PATH, id);
    await deleteDoc(docRef);
    logActivity("delete-template", { templateId: id });
}
async function saveDraftToFirestore(id, data) {
    if (!isFirebaseReady) throw new Error("Firebase não pronto.");
    
    if (!data.data || data.data.some(item => !item.componentLabel)) {
        throw new Error("Draft data is malformed or empty.");
    }
    
    const docRef = id ? doc(db, DRAFTS_COLLECTION_PATH, id) : doc(collection(db, DRAFTS_COLLECTION_PATH));
    await setDoc(docRef, data, { merge: true });
    logActivity("save-draft", { templateId: data.templateId, draftId: docRef.id });
    return docRef.id;
}
async function saveFilledReportToFirestore(draftId, data) {
    if (!isFirebaseReady) throw new Error("Firebase não pronto.");
    
    const docData = { ...data, userId: state.currentUser.id }; 

    const newReportRef = await addDoc(collection(db, FILLED_REPORTS_COLLECTION_PATH), docData);

    if (draftId) {
        const draftRef = doc(db, DRAFTS_COLLECTION_PATH, draftId);
        await deleteDoc(draftRef).catch(e => console.warn("[Firestore] Could not delete draft:", draftId, e));
    }

    logActivity("fill-report", { templateId: data.templateId, filledReportId: newReportRef.id, templateTitle: data.templateTitle });
}
async function deleteDraftFromFirestore(id) {
    if (!isFirebaseReady) throw new Error("Firebase não pronto.");
    const draftRef = doc(db, DRAFTS_COLLECTION_PATH, id);
    await deleteDoc(draftRef);
    logActivity("delete-draft", { draftId: id });
}


// --- Fetchers (Realtime) ---
function fetchReportTemplates() {
    if (!isFirebaseReady) return;
    const reportColRef = collection(db, REPORT_TEMPLATES_COLLECTION_PATH);
    const q = query(reportColRef);
    onSnapshot(q, (snapshot) => {
        state.savedReports = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        renderSavedReports();
    }, (error) => { console.error("[Firestore] Error fetching report templates:", error); showToast("Erro ao carregar modelos de relatório.", "error"); });
}
function fetchDrafts() {
    if (!isFirebaseReady) return;
    const draftsColRef = collection(db, DRAFTS_COLLECTION_PATH);
    const userEmail = state.currentUser?.email;
    if (!userEmail) return;

    const q = query(draftsColRef, where("createdByEmail", "==", userEmail));
    onSnapshot(q, (snapshot) => {
        state.draftReports = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        state.notifications.draftReports = state.draftReports.length;
        updateNotificationBell();
        renderDraftReports();
    }, (error) => { console.error("[Firestore] Error fetching drafts:", error); showToast("Erro ao carregar rascunhos.", "error"); });
}
function fetchFilledReports() {
    if (!isFirebaseReady) return;
    const filledColRef = collection(db, FILLED_REPORTS_COLLECTION_PATH);
    const q = query(filledColRef);
    onSnapshot(q, (snapshot) => {
        state.filledReports = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        renderFilledReports();
    }, (error) => { console.error("[Firestore] Error fetching filled reports:", error); showToast("Erro ao carregar relatórios preenchidos.", "error"); });
}
function fetchUsersRealtime() {
    if (!isFirebaseReady || !state.currentUser?.permissions?.canManageUsers) return;
    const usersColRef = collection(db, USERS_COLLECTION_PATH); 
    onSnapshot(usersColRef, (snapshot) => {
        state.users = []; state.pendingUsers = [];
        snapshot.docs.forEach(doc => {
            const userData = { id: doc.id, ...doc.data() };
            if (userData.autorizado === false) {
                state.pendingUsers.push({...userData, id: doc.id, name: userData.name || 'S/ Nome', email: userData.email || 'S/ Email'});
            }
            state.users.push({...userData, id: doc.id});
        });
        state.notifications.pendingUsers = state.pendingUsers.length;
        updateNotificationBell();
        renderUsersTable(); 
    }, (error) => { console.error("[Firestore] Error fetching users:", error); showToast("Erro ao carregar lista de usuários.", "error"); });
}


// ----------------------------------------------------------------------
// --- FUNÇÕES DE RENDERIZAÇÃO E MANIPULAÇÃO DE DATOS (DOM) -------------
// ----------------------------------------------------------------------

function renderNotificationDropdown() {
    elements.notificationList.innerHTML = ''; let hasNotifications = false;
    if (state.currentUser?.permissions?.canManageUsers && state.notifications.pendingUsers > 0) {
        hasNotifications = true; const count = state.notifications.pendingUsers;
        elements.notificationList.innerHTML += `<div class="notification-item" data-action="goto-users"><div class="icon-wrapper bg-blue-100 text-blue-600"> <i data-lucide="user-plus" class="w-5 h-5"></i> </div><div> <p class="text-sm font-medium">${count} Nova${count > 1 ? 's' : ''} Solicitação${count > 1 ? 's' : ''}</p> <p class="text-xs text-gray-500">Aprove ou negue em 'Usuários'.</p> </div></div>`;
    }
    if (state.notifications.draftReports > 0) {
        hasNotifications = true; const count = state.notifications.draftReports;
        elements.notificationList.innerHTML += `<div class="notification-item" data-action="goto-analysis"><div class="icon-wrapper bg-yellow-100 text-yellow-600"> <i data-lucide="file-clock" class="w-5 h-5"></i> </div><div> <p class="text-sm font-medium">${count} Rascunho${count > 1 ? 's' : ''} Pendente${count > 1 ? 's' : ''}</p> <p class="text-xs text-gray-500">Continue em 'Análise'.</p> </div></div>`;
    }
    if (hasNotifications) { elements.notificationEmptyState.classList.add('hidden'); }
    else { elements.notificationEmptyState.classList.remove('hidden'); }
    try { lucide.createIcons(); } catch(e) { console.error("[Lucide Error] Notification Dropdown:", e); }
}
function handleNotificationClick(e) {
    const item = e.target.closest('.notification-item'); if (!item) return;
    const action = item.dataset.action; console.log("[Notify] Clicked action:", action);
    if (action === 'goto-users') { switchPage('users'); }
    else if (action === 'goto-analysis') { switchPage('analysis'); }
    toggleNotificationDropdown();
}
function toggleNotificationDropdown() {
    elements.notificationDropdown.classList.toggle('active'); elements.notificationDropdown.classList.toggle('hidden');
    if (elements.notificationDropdown.classList.contains('active')) { renderNotificationDropdown(); }
}

function getComponentTypeInfo(type) {
    // Componente Audio removido
    const componentTypes = { 
        text: { name: 'Texto Curto', icon: 'type', options: [] }, 
        textarea: { name: 'Texto Longo', icon: 'align-left', options: [] }, 
        checkbox: { name: 'Seleção', icon: 'check-square', options: ['Item 1', 'Item 2'], singleSelection: false }, 
        select: { name: 'Lista Suspensa', icon: 'list', options: ['Opção 1', 'Opção 2'] }, 
        table: { name: 'Tabela', icon: 'table', options: [], columns: ['Coluna A', 'Coluna B'] }, 
        image: { name: 'Imagem', icon: 'image', options: [] }, 
        date: { name: 'Data', icon: 'calendar', options: [] },
        file: { name: 'Anexar Arquivo', icon: 'paperclip', options: [], allowedTypes: ['xlsx', 'word', 'pdf', 'txt'], maxFileSize: '5MB' }
    };
    return componentTypes[type] || { name: 'Desconhecido', icon: 'help-circle', options: [] };
}
function addComponent(type) {
    const componentId = `comp_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const typeInfo = getComponentTypeInfo(type);
    const component = { id: componentId, type: type, label: typeInfo.name, required: false, options: [...typeInfo.options], columns: typeInfo.columns ? [...typeInfo.columns] : [], singleSelection: typeInfo.singleSelection || false, value: '', width: 'full', order: state.currentReportComponents.length, allowedTypes: typeInfo.allowedTypes || null };
    state.currentReportComponents.push(component);
    renderBuilderComponents(); renderPreviewComponents();
}

function initSortable() { 
    if (elements.reportBuilderArea) {
        new Sortable(elements.reportBuilderArea, {
            animation: 150,
            ghostClass: 'sortable-ghost',
            
            onEnd: function (evt) {
    const { newIndex, oldIndex } = evt;
    if (newIndex === oldIndex) return;
    if (
        oldIndex < 0 ||
        oldIndex >= state.currentReportComponents.length ||
        newIndex < 0 ||
        newIndex > state.currentReportComponents.length
    ) {
        console.warn("[Sortable] Índice inválido detectado:", oldIndex, newIndex);
        return;
    }

    const movedItem = state.currentReportComponents.splice(oldIndex, 1)[0];
    if (!movedItem) {
        console.error("[Sortable] Item movido é indefinido. Abortando reorder.");
        return;
    }

    state.currentReportComponents.splice(newIndex, 0, movedItem);
    state.currentReportComponents.forEach((comp, i) => (comp.order = i));

    renderBuilderComponents();
    renderPreviewComponents();

    console.log(`[Sortable] Reordenado de ${oldIndex} para ${newIndex}.`);
    showToast("Ordem atualizada!");
            }
        });
    }
}


function setupBuilderEventListeners() {
    if (elements.reportBuilderArea) {
        elements.reportBuilderArea.addEventListener('input', (event) => {
             const target = event.target;
             if (target.classList.contains('component-label-input')) {
                 const card = target.closest('.component-card');
                 const componentId = card?.dataset.id;
                 if (componentId) {
                     updateComponentLabel(componentId, target.value);
                 }
             }
         });
        elements.reportBuilderArea.addEventListener('click', (event) => {
             const button = event.target.closest('button');
             if (!button) return;
             const card = button.closest('.component-card');
             const componentId = card?.dataset.id;
             const action = button.dataset.action;
             if (componentId && action) {
                 switch (action) {
                     case 'edit': editComponent(componentId); break;
                     case 'delete': deleteComponent(componentId); break;
                     case 'toggle-width': toggleComponentWidth(componentId); break;
                 }
             }
         });
    }
    document.querySelectorAll('.component-type').forEach(component => {
        component.addEventListener('click', () => {
            const type = component.dataset.type;
            addComponent(type);
        });
    });
}
function updateComponentLabel(id, newLabel) {
    const component = state.currentReportComponents.find(comp => comp.id === id);
    if (component) { component.label = newLabel || 'S/ Rótulo'; renderPreviewComponents(); }
}
function toggleComponentWidth(id) {
    const component = state.currentReportComponents.find(comp => comp.id === id);
    if (component) { component.width = component.width === 'full' ? 'half' : 'full'; renderBuilderComponents(); renderPreviewComponents(); }
}
function editComponent(id) {
    const component = state.currentReportComponents.find(comp => comp.id === id); if (!component) { return; }
    state.editingComponentId = id; elements.componentModalTitle.textContent = `Editar ${getComponentTypeInfo(component.type).name}`; elements.componentId.value = id; elements.componentLabel.value = component.label; elements.componentRequired.checked = component.required;
    const showOptions = component.type === 'select' || component.type === 'checkbox'; 
    const showColumns = component.type === 'table'; 
    const showSingleSelection = component.type === 'checkbox';
    
    elements.componentOptionsContainer.classList.toggle('hidden', !showOptions); 
    elements.componentColumnsContainer.classList.toggle('hidden', !showColumns); 
    elements.componentSingleSelectionContainer.classList.toggle('hidden', !showSingleSelection);
    
    if (showOptions) elements.componentOptions.value = (component.options || []).join('\n'); 
    if (showColumns) elements.componentColumns.value = (component.columns || []).join('\n'); 
    if (showSingleSelection) elements.componentSingleSelection.checked = component.singleSelection;
    
    elements.componentModal.classList.add('active'); elements.componentLabel.focus(); elements.componentLabel.select();
}
function deleteComponent(id) {
    const component = state.currentReportComponents.find(comp => comp.id === id); const label = component ? component.label : 'este componente';
    showConfirmModal("Excluir Componente", `Remover "${label}"?`, () => {
        state.currentReportComponents = state.currentReportComponents.filter(comp => comp.id !== id); state.currentReportComponents.forEach((comp, index) => { comp.order = index; }); renderBuilderComponents(); renderPreviewComponents(); showToast('Componente removido.'); });
}
function handleComponentFormSubmit(e) {
    e.preventDefault(); const id = elements.componentId.value; const component = state.currentReportComponents.find(comp => comp.id === id);
    if (component) {
        component.label = elements.componentLabel.value.trim() || `Campo ${component.type}`; component.required = elements.componentRequired.checked;
        if (component.type === 'select' || component.type === 'checkbox') { component.options = elements.componentOptions.value.split('\n').map(opt=>opt.trim()).filter(opt=>opt !== ''); if(component.options.length===0) component.options = ['Opção Padrão']; }
        if (component.type === 'table') { component.columns = elements.componentColumns.value.split('\n').map(col=>col.trim()).filter(col=>col !== ''); if(component.columns.length===0) component.columns = ['Coluna Padrão']; }
        if (component.type === 'checkbox') component.singleSelection = elements.componentSingleSelection.checked;
        renderBuilderComponents(); renderPreviewComponents(); showToast('Componente atualizado!');
    } else { console.error("[Builder] Component not found during form submit:", id); }
    closeComponentModal();
}
function saveReport() {
    const title = elements.reportTitle.value.trim(); const editingId = state.editingReportId;
    if (!title) { showToast('Título obrigatório.', 'error'); elements.reportTitle.focus(); return; } if (state.currentReportComponents.length === 0) { showToast('Adicione componentes.', 'error'); return; }
    const componentsToSave = [...state.currentReportComponents].sort((a,b)=>a.order-b.order).map(({ id, ...rest }) => ({...rest, order: rest.order})); 
    if (!state.currentUser) { showToast('Usuário não autenticado.', 'error'); return; }

    const reportData = { title: title, components: componentsToSave, updatedBy: state.currentUser.email, lastUpdatedAt: new Date().toISOString(), status: 'active' };
    if (!editingId) { reportData.createdBy = state.currentUser.email; reportData.createdAt = new Date().toISOString(); }

    if (typeof saveReportTemplateToFirestore === 'function' && state.currentUser.id !== 'local_admin') {
        const originalText = elements.saveReportBtn.innerHTML;
        elements.saveReportBtn.innerHTML = '<i data-lucide="loader" class="w-5 h-5 animate-spin"></i> Salvando...';
        elements.saveReportBtn.disabled = true;
        
        saveReportTemplateToFirestore(editingId, reportData)
            .then(() => { showToast(`Relatório ${editingId ? 'atualizado' : 'salvo'} com sucesso!`); clearReportBuilder(); switchPage('fill-report'); })
            .catch((error) => { console.error("[Builder] Erro ao salvar no Firestore:", error); showToast('Erro ao salvar no Firestore. Verifique as regras de segurança.', 'error'); })
            .finally(() => { elements.saveReportBtn.innerHTML = originalText; elements.saveReportBtn.disabled = false; try { lucide.createIcons(); } catch(e) {} });
    } else {
        if (editingId) {
            const reportIndex = state.savedReports.findIndex(r => r.id === editingId);
            if (reportIndex > -1) { state.savedReports[reportIndex] = { ...state.savedReports[reportIndex], ...reportData, id: editingId }; showToast('Relatório atualizado (Local)!'); } 
        } else {
            const newReport = { id: `report_${Date.now()}`, ...reportData, createdBy: state.currentUser?.name || 'Admin', status: 'active' };
            state.savedReports.push(newReport);
            showToast('Relatório salvo (Local)!');
        }
        clearReportBuilder(); renderSavedReports();
        switchPage('fill-report');
    }
}
function loadReportForEditing(reportId) {
    const report = state.savedReports.find(r => r.id === reportId);
    if (report) {
        state.editingReportId = reportId; elements.editingReportId.value = reportId; elements.reportTitle.value = report.title;
        state.currentReportComponents = (report.components || []).map((comp, index) => ({ ...comp, id: `comp_${Date.now()}_${index}`, order: comp.order ?? index })).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
        renderBuilderComponents(); renderPreviewComponents(); updatePreviewTitle();
        elements.saveReportBtn.innerHTML = '<i data-lucide="save" class="w-5 h-5"></i> Atualizar Relatório';
        try { lucide.createIcons({nodes:[elements.saveReportBtn.querySelector('i')]}); } catch(e) {}
        elements.cancelEditBtn.classList.remove('hidden'); switchPage('create-report');
    } else { showToast('Relatório não encontrado.', 'error'); }
}
function cancelReportEdit(doSwitchPage = true) {
    clearReportBuilder();
    if (doSwitchPage) { switchPage('fill-report'); }
}
function clearReportBuilder() {
    state.currentReportComponents = []; state.editingReportId = null; elements.reportTitle.value = ''; elements.editingReportId.value = ''; elements.saveReportBtn.innerHTML = '<i data-lucide="save" class="w-5 h-5"></i> Salvar Relatório'; elements.cancelEditBtn.classList.add('hidden');
    try { lucide.createIcons({nodes:[elements.saveReportBtn.querySelector('i')]}); } catch(e) {}
    renderBuilderComponents(); renderPreviewComponents(); updatePreviewTitle();
}
function deleteReport(reportId) {
    const report = state.savedReports.find(r => r.id === reportId);
    showConfirmModal("Excluir Modelo", `Excluir "${report?.title || reportId}"?`, () => {
        if (typeof deleteReportTemplateFromFirestore === 'function' && state.currentUser.id !== 'local_admin') {
            deleteReportTemplateFromFirestore(reportId)
               .then(() => showToast('Modelo excluído!'))
               .catch(() => showToast('Erro ao excluir no Firestore.', 'error'));
        } else {
            state.savedReports = state.savedReports.filter(r => r.id !== reportId); renderSavedReports(); showToast('Modelo excluído (Local).');
        }
    });
}
function toggleReportStatus(reportId) {
    const report = state.savedReports.find(r => r.id === reportId); if (!report) return;
    const newStatus = report.status === 'active' ? 'paused' : 'active';
     if (typeof updateReportTemplateStatusInFirestore === 'function' && state.currentUser.id !== 'local_admin') {
          updateReportTemplateStatusInFirestore(reportId, newStatus)
             .then(() => showToast(`Status alterado para ${newStatus}.`))
             .catch(() => showToast('Erro ao atualizar status.', 'error'));
     } else {
       report.status = newStatus; renderSavedReports(); showToast(`Status alterado para ${newStatus} (Local).`);
     }
}
function getFilledData(form) {
    const reportTemplate = state.savedReports.find(r => r.id === currentFillReportId); if (!reportTemplate) { console.error("[Fill] Template not found."); return null; }
    const filledData = [];
    
    (reportTemplate.components || []).forEach((component, index) => { 
        const inputNameBase = `component_${component.order ?? index}`; 
        let value = null;

        // Tratamento de arquivos e imagens (que usam input hidden para o valor final)
        if (component.type === 'image' || component.type === 'file') {
             const hiddenInput = form.querySelector(`input[name="${inputNameBase}"]`);
             // Limpar o valor do file input para não enviar dados binários (causa erro 400)
             const fileInput = form.querySelector(`input[id="${'fill'}_${component.id}"]`);
             if (fileInput) fileInput.value = '';
             
             value = hiddenInput?.value || null;
        }
        
        // Tratamento de checkbox
        else if (component.type === 'checkbox' && !component.singleSelection) { 
            value = Array.from(form.querySelectorAll(`input[name^="${inputNameBase}_"]:checked`)).map(el => el.value); 
        }
        
        // Tratamento de tabela
        else if (component.type === 'table') {
            value = []; const tableContainer = form.querySelector(`[data-input-name="${inputNameBase}"]`);
            tableContainer?.querySelectorAll('.dynamic-table-row').forEach(rowEl => { 
                const rowObject = {}; let rowHasValue = false;
                rowEl.querySelectorAll('input[type="text"]').forEach(inputEl => {
                    const colName = inputEl.dataset.colName; rowObject[colName] = inputEl.value; if(inputEl.value) rowHasValue = true;
                });
                if(rowHasValue) value.push(rowObject);
            });
        }
        
        // Tratamento de outros inputs simples (text, textarea, select, date, radio, audio)
        else { 
            const input = form.querySelector(`[name="${inputNameBase}"]`);
            value = input ? input.value : null; 
        }

        filledData.push({ componentLabel: component.label, componentType: component.type, value: value ?? '' });
    });
    
    // Filtra dados com valor nulo para tentar evitar erros de serialização no Firestore
    const cleanedFilledData = filledData.filter(item => item.value !== null && item.value !== undefined);
    return cleanedFilledData;
}

function handleSaveDraft(e) {
    e.preventDefault();
    if (!currentFillReportId || !state.currentUser) { showToast('Usuário ou relatório não disponível.', 'error'); return; }
    const reportTemplate = state.savedReports.find(r => r.id === currentFillReportId); if (!reportTemplate) return;
    
    // Coletar apenas dados serializáveis
    const filledData = getFilledData(elements.fillReportForm); 
    if (!filledData || filledData.length === 0) { 
        showToast("Nenhum dado válido para salvar no rascunho.", "warning"); return; 
    }
    
    const existingDraftId = elements.fillReportForm.dataset.draftId;

    const draftData = { 
        templateId: currentFillReportId, 
        templateTitle: reportTemplate.title, 
        createdByEmail: state.currentUser.email, 
        createdByName: state.currentUser.name, 
        lastUpdatedAt: new Date().toISOString(), 
        data: filledData
    };
    
    if (typeof saveDraftToFirestore === 'function' && state.currentUser.id !== 'local_admin') {
        const draftButton = elements.fillReportForm.querySelector('#save-draft-btn-modal');
        const originalText = draftButton.innerHTML;
        draftButton.innerHTML = '<i data-lucide="loader" class="w-5 h-5 animate-spin"></i> Salvando...';
        draftButton.disabled = true;

        saveDraftToFirestore(existingDraftId, draftData)
           .then(() => { showToast("Rascunho salvo!"); closeFillReportModal(); })
           .catch((error) => { 
               console.error('[Firestore] Erro ao salvar rascunho:', error);
               showToast('Erro ao salvar rascunho (Verifique regras do Firestore).', 'error');
           })
           .finally(() => { 
               draftButton.innerHTML = originalText; 
               draftButton.disabled = false; 
               try { lucide.createIcons(); } catch(e) {} 
           });
    } else {
        if (existingDraftId) {
            const draftIndex = state.draftReports.findIndex(d => d.id === existingDraftId);
            if (draftIndex > -1) { state.draftReports[draftIndex].data = filledData; state.draftReports[draftIndex].lastUpdatedAt = new Date().toISOString(); }
        } else {
            const draftReport = { id: `draft_${Date.now()}`, ...draftData, createdAt: new Date().toISOString() };
            state.draftReports.push(draftReport);
        }
        state.notifications.draftReports = state.notifications.draftReports - 1;
        updateNotificationBell(); renderDraftReports(); showToast("Rascunho salvo (Local)!"); closeFillReportModal();
    }
}
function handleFillReportSubmit(e) {
    e.preventDefault();
    if (!currentFillReportId || !state.currentUser) { showToast('Usuário ou relatório não disponível.', 'error'); return; }
    const reportTemplate = state.savedReports.find(r => r.id === currentFillReportId); if (!reportTemplate) return;
    
    const filledData = getFilledData(elements.fillReportForm); 
    if (!filledData || filledData.length === 0) {
        showToast("Por favor, preencha o formulário antes de enviar.", "warning"); return;
    }
    
    let isValid = true;
    (reportTemplate.components || []).forEach((component) => {
        if (component.required) {
            const filledComponent = filledData.find(d => d.componentLabel === component.label && d.componentType === component.type);
            if (!filledComponent || filledComponent.value === null || filledComponent.value === undefined || (typeof filledComponent.value === 'string' && filledComponent.value.trim() === '') || (Array.isArray(filledComponent.value) && filledComponent.value.length === 0)) {
                isValid = false;
            }
        }
    });
    if (!isValid) {
        showToast("Por favor, preencha todos os campos obrigatórios.", 'error');
        return;
    }

    const filledReport = { 
        templateId: currentFillReportId, 
        templateTitle: reportTemplate.title, 
        filledAt: new Date().toISOString(), 
        filledByEmail: state.currentUser.email, 
        filledByName: state.currentUser.name, 
        data: filledData
    };
    const existingDraftId = elements.fillReportForm.dataset.draftId;

    if (typeof saveFilledReportToFirestore === 'function' && state.currentUser.id !== 'local_admin') {
        const submitButton = elements.fillReportForm.querySelector('button[type="submit"]');
        const originalText = submitButton.innerHTML;
        submitButton.innerHTML = '<i data-lucide="loader" class="w-5 h-5 animate-spin"></i> Salvando...';
        submitButton.disabled = true;

        saveFilledReportToFirestore(existingDraftId, filledReport)
             .then(() => { showToast("Salvo com sucesso!"); closeFillReportModal(); })
             .catch((error) => { 
                 console.error('[Firestore] Erro ao enviar relatório:', error);
                 showToast('Erro ao enviar relatório (Verifique regras do Firestore).', 'error');
             })
             .finally(() => { submitButton.innerHTML = originalText; submitButton.disabled = false; try { lucide.createIcons(); } catch(e) {} });
    } else {
        filledReport.id = `filled_${Date.now()}`;
        state.filledReports.push(filledReport);
        if (existingDraftId) {
            state.draftReports = state.draftReports.filter(d => d.id !== existingDraftId); state.notifications.draftReports = state.notifications.draftReports - 1;
            updateNotificationBell(); renderDraftReports();
        }
        showToast("Salvo com sucesso (Local)!"); closeFillReportModal(); renderFilledReports();
    }
}
function handleTableAddRow(buttonEl) {
    const tableContainer = buttonEl.closest('[data-component-type="table"]');
    if (!tableContainer) { return; }
    const rowsContainer = tableContainer.querySelector('.dynamic-table-rows-container');
    const firstRow = rowsContainer.querySelector('.dynamic-table-row');
    
    if (firstRow) {
        const newRow = firstRow.cloneNode(true);
        newRow.querySelectorAll('input').forEach(input => {
            input.value = ''; 
            const name = input.getAttribute('name'); 
            const newName = name.replace(/_\d+_/, `_${Date.now()}_`); 
            input.setAttribute('name', newName);
        });
        
        let removeButton = newRow.querySelector('.table-remove-row');
        if (!removeButton) {
            // Se a linha original não tinha botão (porque era a única), adiciona um.
            newRow.insertAdjacentHTML('beforeend', `<button type="button" class="btn btn-sm btn-danger table-remove-row" title="Remover Linha" onclick="handleTableRemoveRow(this)"><i data-lucide="minus"></i></button>`);
        } else {
            removeButton.setAttribute('onclick', 'handleTableRemoveRow(this)');
        }
        
        // Adiciona um botão de remoção à primeira linha se ela não tiver (a partir da segunda linha adicionada)
        if (rowsContainer.children.length === 1 && !rowsContainer.querySelector('.dynamic-table-row .table-remove-row')) {
            rowsContainer.querySelector('.dynamic-table-row').insertAdjacentHTML('beforeend', `<button type="button" class="btn btn-sm btn-danger table-remove-row" title="Remover Linha" onclick="handleTableRemoveRow(this)"><i data-lucide="minus"></i></button>`);
        }
        
        rowsContainer.appendChild(newRow);
        try { lucide.createIcons(); } catch(e) {}
    } else { showToast("Erro ao adicionar linha: Tabela vazia.", "error"); }
}
function handleTableRemoveRow(buttonEl) {
    const row = buttonEl.closest('.dynamic-table-row');
    const rowsContainer = row.parentElement;
    
    if (rowsContainer.children.length > 1) { 
        row.remove();
        // Remove o botão de remoção da última linha se ela se tornar a única
        if (rowsContainer.children.length === 1) {
            rowsContainer.querySelector('.table-remove-row')?.remove();
        }
    } else {
        showToast("A tabela precisa de pelo menos uma linha.", "warning");
    }
}
function handleUserFormSubmit(e) {
    e.preventDefault();
    const name = elements.userForm.querySelector('#user-name').value.trim();
    const email = elements.userForm.querySelector('#user-email').value.trim();
    const profile = elements.userForm.querySelector('#user-profile').value;
    const userId = elements.userForm.querySelector('#user-id').value;
    const password = elements.userPasswordInput?.value;
    const permissions = {}; elements.userForm.querySelectorAll('input[type="checkbox"][data-permission]').forEach(check => { permissions[check.dataset.permission] = check.checked; });

    if (!name || !email || !/\S+@\S+\.\S+/.test(email)) { showToast('Nome e Email válidos são obrigatórios', 'error'); return; }

    const submitBtn = elements.userForm.querySelector('button[type="submit"]');
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<i data-lucide="loader" class="w-5 h-5 animate-spin"></i> Salvando...';
    try { lucide.createIcons(); } catch(e) {}

    const saveUser = async (uid) => {
          const newUserRecord = { name, email, profile, autorizado: true, createdAt: new Date().toISOString(), permissions };
          const userRef = doc(db, USERS_COLLECTION_PATH, uid);
          await setDoc(userRef, newUserRecord);
          showToast('Usuário adicionado e acesso garantido!');
          logActivity("create-user", { targetUid: uid, targetEmail: email, profile });
          closeUserModal();
    };

    try {
        if (!userId) { 
            if (elements.userPasswordInput && (!password || password.length < 6)) { showToast('A senha é obrigatória e deve ter no mínimo 6 caracteres para novos usuários.', 'error'); submitBtn.disabled = false; submitBtn.innerHTML = 'Salvar'; try { lucide.createIcons(); } catch(e) {} return; }
            if (isFirebaseReady) {
                createUserWithEmailAndPassword(auth, email, password)
                    .then((userCredential) => { saveUser(userCredential.user.uid); })
                    .catch((error) => {
                        let msg = "Erro ao adicionar usuário.";
                        if (error.code === 'auth/email-already-in-use') { msg = 'Este email já está em uso.'; }
                        else if (error.code === 'auth/weak-password') { msg = 'A senha é muito fraca.'; }
                        showToast(msg, 'error');
                    });
            } else { // Fallback Local
                if (state.users.some(u => u.email === email)) { showToast('Email já cadastrado (local).', 'error'); }
                else {
                    const newUser = { id: `local_${Date.now()}`, name, email, profile, permissions, autorizado: true };
                    state.users.push(newUser); showToast('Usuário adicionado (Local)!'); renderUsersTable(); closeUserModal();
                }
            }
        } else { 
            if (isFirebaseReady && userId !== 'local_admin') {
                const userRef = doc(db, USERS_COLLECTION_PATH, userId);
                const updateData = { name, email, profile, permissions };
                updateDoc(userRef, updateData).then(() => {
                    showToast('Usuário atualizado!');
                    logActivity("update-user", { targetUid: userId, targetEmail: email, profile });
                }).catch(() => { showToast("Erro ao atualizar usuário no Firestore.", 'error'); });
            } else { // Edição Local
                const userIndex = state.users.findIndex(u => u.id == userId);
                if (userIndex > -1) { state.users[userIndex] = { ...state.users[userIndex], name, email, profile, permissions }; showToast('Usuário atualizado (Local)!'); renderUsersTable(); }
                else { showToast('Erro ao editar (Local).', 'error'); }
            }
            closeUserModal();
        }
    } finally {
        setTimeout(() => {
            submitBtn.disabled = false;
            submitBtn.innerHTML = 'Salvar';
            try { lucide.createIcons(); } catch(e) {}
        }, 500);
    }
}
function deleteUser(id) {
    const user = state.users.find(u => u.id == id);
    if(user && user.profile === 'admin' && user.id === state.currentUser?.id) { showToast("Você não pode se excluir.", "warning"); return;}

    if (user) {
        showConfirmModal("Excluir Usuário", `Excluir ${user.name}?`, async () => {
            if (isFirebaseReady && user.id !== 'local_admin') {
                try {
                    await deleteDoc(doc(db, USERS_COLLECTION_PATH, user.id));
                    showToast('Usuário excluído (Firestore)!');
                    logActivity("delete-user", { targetUid: user.id, targetEmail: user.email });
                    if (user.id === auth.currentUser?.uid) {
                         signOut(auth);
                         showToast('Sua conta foi excluída.', 'warning');
                    }
                } catch (e) { 
                    showToast("Erro ao excluir registro do usuário. Verifique as regras de segurança.", 'error'); 
                    console.error("[Firestore] Delete User Error:", e);
                }
            } else {
                state.users = state.users.filter(u => u.id != id);
                renderUsersTable(); showToast('Usuário excluído (Local).');
            }
        });
    }
}

function renderBuilderComponents() {
    if (!elements.reportBuilderArea) { return; }
    if (state.currentReportComponents.length === 0) {
        elements.reportBuilderArea.innerHTML = `<div class="text-center py-10 text-gray-500 dark:text-gray-400"><i data-lucide="layout-template" class="w-12 h-12 mx-auto mb-3 opacity-50"></i><p>Clique ou arraste componentes</p></div>`;
        try { lucide.createIcons(); } catch(e) {}
        return;
    }
    
    const sortedComponents = [...state.currentReportComponents].sort((a, b) => a.order - b.order);
    let html = '';

    for (let i = 0; i < sortedComponents.length; i++) {
        const comp = sortedComponents[i];
        const typeInfo = getComponentTypeInfo(comp.type);
        const widthIcon = comp.width === 'full' ? 'columns' : 'maximize-2';
        const widthTitle = comp.width === 'full' ? 'Meia Largura' : 'Largura Total';
        
        const componentHtml = `<div class="component-card ${comp.width==='half'?'w-1/2 float-left':'w-full'}" data-id="${comp.id}">
            <div class="flex items-center justify-between mb-2 gap-2">
                <div class="flex items-center gap-2 flex-grow min-w-0">
                    <i data-lucide="${typeInfo.icon}" class="w-5 h-5 text-blue-500 flex-shrink-0"></i>
                    <input type="text" class="component-label-input flex-grow min-w-0" value="${comp.label}" placeholder="Rótulo">
                </div>
                <div class="flex items-center gap-1 flex-shrink-0">
                    <button class="btn btn-sm btn-secondary p-1" data-action="toggle-width" title="${widthTitle}"><i data-lucide="${widthIcon}" class="w-4 h-4"></i></button>
                    <button class="btn btn-sm btn-secondary p-1" data-action="edit" title="Editar"><i data-lucide="settings" class="w-4 h-4"></i></button>
                    <button class="btn btn-sm btn-danger p-1" data-action="delete" title="Excluir"><i data-lucide="trash-2" class="w-4 h-4"></i></button>
                </div>
            </div>
            <div class="text-xs text-gray-500 flex items-center gap-2">
                <span>${typeInfo.name}</span>
                ${comp.required ? '<span class="text-red-500">*</span>' : ''}
                ${comp.width==='half'?'<span class="text-indigo-500">(1/2)</span>':''}
                ${comp.type==='checkbox'&&comp.singleSelection?'<span class="text-purple-500">(Única)</span>':''}
            </div>
        </div>`;
        
        html += componentHtml;

        if (comp.width === 'full') {
            html += '<div class="clearfix"></div>';
        } else if (comp.width === 'half') {
            const nextComp = sortedComponents[i + 1];
            
            if (!nextComp || nextComp.width !== 'half') {
                html += '<div class="clearfix"></div>';
            }
        }
    }
    
    elements.reportBuilderArea.innerHTML = html;
    try { lucide.createIcons(); } catch(e) {}
}

function renderPreviewComponents() {
    if (!elements.previewComponents) return;
    if (state.currentReportComponents.length === 0) { elements.previewComponents.innerHTML = `<p class="text-gray-500 dark:text-gray-400 text-center py-10">Preview vazio</p>`; return; }
    let html = ''; const sortedComponents = [...state.currentReportComponents].sort((a, b) => a.order - b.order);
    for (let i = 0; i < sortedComponents.length; i++) { const comp = sortedComponents[i]; const nextComp = sortedComponents[i + 1]; const isHalf = comp.width === 'half'; const nextIsHalf = nextComp?.width === 'half'; if (isHalf && nextIsHalf) { html += `<div class="grid grid-cols-2 gap-4"><div>${renderSinglePreviewComponent(comp)}</div><div>${renderSinglePreviewComponent(nextComp)}</div></div>`; i++; } else { html += renderSinglePreviewComponent(comp); } }
    elements.previewComponents.innerHTML = html;
    try { lucide.createIcons(); } catch(e) {}
}
function renderSinglePreviewComponent(component, isFillForm = false, value = null) {
     const inputId = `${isFillForm ? 'fill' : 'preview'}_${component.id}`; 
     const nameAttr = `component_${component.order}`;
     const requiredAttr = component.required && isFillForm ? 'required' : '';
     const baseClass = isFillForm ? 'form-input' : 'form-input pointer-events-none opacity-70';

     let inputHtml = '';
     switch (component.type) {
         case 'text': 
             // Não adicionamos botão de ditado aqui
             inputHtml = `<input type="text" id="${inputId}" name="${nameAttr}" class="${baseClass}" placeholder="Digite..." value="${value || ''}" ${requiredAttr}>`;
             break;
         case 'textarea': 
             // Ditado é exclusivo do textarea
             inputHtml = `<div class="relative"><textarea id="${inputId}" name="${nameAttr}" class="${baseClass} pr-10" rows="3" placeholder="Digite ou dite..." ${requiredAttr}>${value || ''}</textarea><button type="button" class="dictation-button absolute right-3 top-3" data-target="${inputId}" title="Ditar" onclick="handleDictationClick(this)"><i data-lucide="mic" class="w-5 h-5"></i></button></div>`;
             break;
         case 'checkbox':
             if (component.singleSelection) {
                 inputHtml = `<div class="space-y-2">${(component.options || []).map((opt, index) => `<label class="flex items-center gap-2 text-sm"><input type="radio" id="${inputId}_${index}" name="${nameAttr}" value="${opt}" class="rounded-full text-blue-500" ${requiredAttr} ${value === opt ? 'checked' : ''}><span>${opt}</span></label>`).join('')}</div>`;
             } else {
                 const values = Array.isArray(value) ? value : [];
                 inputHtml = `<div class="space-y-2">${(component.options || []).map((opt, index) => `<label class="flex items-center gap-2 text-sm"><input type="checkbox" id="${inputId}_${index}" name="${nameAttr}_${index}" value="${opt}" class="rounded text-blue-500" ${values.includes(opt) ? 'checked' : ''}><span>${opt}</span></label>`).join('')}</div>`;
             }
             break;
         case 'select':
             inputHtml = `<select id="${inputId}" name="${nameAttr}" class="${baseClass}" ${requiredAttr}><option value="">-- Selecione --</option>${(component.options || []).map(opt => `<option value="${opt}" ${value === opt ? 'selected' : ''}>${opt}</option>`).join('')}</select>`;
             break;
         case 'table':
             const cols = component.columns || [];
             inputHtml = `<div class="border rounded-lg overflow-hidden p-4" data-component-type="table" data-input-name="${nameAttr}"><div class="dynamic-table-header" style="grid-template-columns: repeat(${cols.length}, 1fr) 40px;">${cols.map(col => `<span>${col}</span>`).join('')}<span></span></div><div class="dynamic-table-rows-container space-y-2">`;
             
             let currentRows = Array.isArray(value) && value.length > 0 ? value : [{}];
             
             inputHtml += currentRows.map((rowData, rowIndex) => {
                 const removeBtn = rowIndex > 0 || currentRows.length > 1 ? `<button type="button" class="btn btn-sm btn-danger table-remove-row" title="Remover Linha" onclick="handleTableRemoveRow(this)"><i data-lucide="minus"></i></button>` : `<span class="w-10"></span>`;

                 return `<div class="dynamic-table-row" style="grid-template-columns: repeat(${cols.length}, 1fr) 40px;"><div class="dynamic-table-row-inputs" style="grid-column: 1 / span ${cols.length}; grid-template-columns: repeat(${cols.length}, 1fr);">${cols.map((col, colIndex) => `<input type="text" name="${nameAttr}_${rowIndex}_${colIndex}" data-col-name="${col}" class="${baseClass} form-input-sm" placeholder="..." value="${rowData[col] || ''}">`).join('')}</div>${removeBtn}</div>`;
             }).join('');
             
             if (isFillForm) {
                 inputHtml += `</div><button type="button" class="btn btn-sm btn-secondary mt-3" onclick="handleTableAddRow(this)"><i data-lucide="plus" class="w-4 h-4"></i> Adicionar Linha</button></div>`;
             } else {
                 inputHtml += `</div></div>`;
             }
             break;
         case 'audio': 
             // COMPONENTE AUDIO (Revertido para um campo simples, sem ditado próprio)
             inputHtml = `<div class="relative flex items-center gap-3 p-3 bg-gray-100 dark:bg-gray-800 rounded-lg">
                           ${isFillForm ? `<input type="text" id="${inputId}" class="${baseClass} flex-1" placeholder="Gravação de áudio desabilitada. Use o Texto Longo para ditado." value="${value || ''}" name="${nameAttr}" disabled>
                                          <div class="text-xs text-gray-500 absolute bottom-0 left-3">[Recurso de Gravação de Áudio]</div>` 
                           : `<div class="flex-1 p-1 text-sm text-gray-500">[Gravação de Áudio]</div>`}
                         </div>`;
             break;
             
         case 'image': 
             const imgPreview = value ? `<p class="text-xs text-green-600 dark:text-green-400">Arquivo: ${value}</p>` : 'Nenhum arquivo anexado.';
             const acceptImage = 'image/*';
             inputHtml = `<div class="flex flex-col gap-2 p-3 bg-gray-100 dark:bg-gray-800 rounded-lg" data-comp-type="image">
                          ${isFillForm ? `<input type="file" id="${inputId}" class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-yellow-50 file:text-yellow-700 hover:file:bg-yellow-100" accept="${acceptImage}" ${requiredAttr} onchange="handleFileUpload(this, 'image')">
                                         <input type="hidden" id="${inputId}_value" name="${nameAttr}" value="${value || ''}">
                                         <div id="${inputId}_preview" class="text-xs text-gray-500">${imgPreview}</div>` : `<div class="flex-1 p-1 text-sm text-gray-500">[Captura/Upload de Imagem]</div>`}
                         </div>`;
             break;
             
         case 'file':
             const allowedExts = (component.allowedTypes || ['pdf', 'txt']).join(', ');
             const acceptFile = (component.allowedTypes || []).map(ext => {
                 if (ext === 'word') return '.doc,.docx';
                 if (ext === 'xlsx') return '.xlsx';
                 return `.${ext}`;
             }).join(',');
             const filePreview = value ? `<p class="text-xs text-green-600 dark:text-green-400">Arquivo: ${value}</p>` : 'Nenhum arquivo anexado.';

             inputHtml = `<div class="flex flex-col gap-2 p-3 bg-gray-100 dark:bg-gray-800 rounded-lg" data-comp-type="file">
                          ${isFillForm ? `<input type="file" id="${inputId}" class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-orange-50 file:text-orange-700 hover:file:bg-orange-100" accept="${acceptFile}" ${requiredAttr} onchange="handleFileUpload(this, 'file')">
                                         <input type="hidden" id="${inputId}_value" name="${nameAttr}" value="${value || ''}">
                                         <div id="${inputId}_preview" class="text-xs text-gray-500">Aceita: ${allowedExts.toUpperCase()}. ${filePreview}</div>` : `<div class="flex-1 p-1 text-sm text-gray-500">[Anexo de Arquivo]</div>`}
                         </div>`;
             break;

         case 'date': 
             inputHtml = `<input type="date" id="${inputId}" name="${nameAttr}" class="${baseClass}" value="${value || ''}" ${requiredAttr}>`;
             break;
         default: 
             inputHtml = `<input type="text" id="${inputId}" name="${nameAttr}" class="${baseClass}" placeholder="Campo" value="${value || ''}" ${requiredAttr}>`;
             break;
     }

     return `<div class="space-y-2"><label class="block text-sm font-medium dark:text-gray-300">${component.label}${component.required ? '<span class="text-red-500">*</span>' : ''}</label>${inputHtml}</div>`;
}

// --- FUNÇÃO DE MANIPULAÇÃO E COMPACTAÇÃO DE ARQUIVOS (Simulada) ---
function handleFileUpload(inputElement, type) {
    const file = inputElement.files[0];
    const inputId = inputElement.id;
    const hiddenValueInput = document.getElementById(inputId + '_value');
    const previewText = document.getElementById(inputId + '_preview');

    if (!file) {
        hiddenValueInput.value = '';
        previewText.innerHTML = 'Nenhum arquivo anexado.';
        return;
    }

    const originalSizeKB = (file.size / 1024).toFixed(2);
    
    hiddenValueInput.value = '';
    previewText.innerHTML = `<i data-lucide="loader" class="w-4 h-4 animate-spin inline mr-2"></i> Processando ${type}... (${originalSizeKB} KB)`;
    try { lucide.createIcons(); } catch(e) {}
    
    // Simulação de Compactação (Real ou Conceitual)
    if (type === 'image' && window.webkitSpeechRecognition) { // Usamos a verificação da API para simular um navegador moderno
        const reader = new FileReader();
        reader.onload = (event) => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width; canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, img.width, img.height);
                
                canvas.toBlob((blob) => {
                    const compressedSizeKB = (blob.size / 1024).toFixed(2);
                    const newFileName = `${file.name.split('.')[0]}_COMP.jpg`;
                    
                    hiddenValueInput.value = newFileName;
                    showToast(`Imagem compactada de ${originalSizeKB} KB para ${compressedSizeKB} KB.`, 'success');
                    previewText.innerHTML = `Arquivo: ${newFileName} (Economia de espaço!)`;
                    inputElement.value = ''; // Limpar o file input para evitar serialização binária
                }, 'image/jpeg', 0.8);
            };
            img.src = event.target.result;
        };
        reader.readAsDataURL(file);
    } 
    else {
        setTimeout(() => {
            const fileNameParts = file.name.split('.');
            const fileExtension = fileNameParts.length > 1 ? fileNameParts.pop() : '';
            const fileNameBase = fileNameParts.join('.');
            const compressedFileName = `${fileNameBase}_COMP.${fileExtension}`;
            
            hiddenValueInput.value = compressedFileName;
            showToast(`Arquivo "${file.name}" pronto para upload (compactação simulada).`, 'success');
            previewText.innerHTML = `Arquivo: ${compressedFileName} (Compactação simulada)`;
            inputElement.value = ''; // Limpar o file input para evitar serialização binária
        }, 1000); 
    }
}


function fillReport(reportId, draftId = null) {
    const report = state.savedReports.find(r => r.id === reportId);
    if (!report || report.status === 'paused') { showToast("Relatório não disponível.", "error"); return; }
    if (!state.currentUser?.permissions.canFillReports) { showToast("Você não tem permissão para preencher relatórios.", "error"); return; }

    currentFillReportId = reportId;
    elements.fillReportModalTitle.textContent = `Preencher Relatório: ${report.title}`;
    
    let draftData = {};
    if (draftId) {
        const draft = state.draftReports.find(d => d.id === draftId);
        if (draft) {
            report.components.forEach(templateComp => {
                const filledItem = draft.data.find(item => item.componentLabel === templateComp.label && item.componentType === templateComp.type);
                if (filledItem) {
                    draftData[templateComp.order] = filledItem.value;
                }
            });
            showToast("Rascunho carregado. Continue preenchendo.");
        }
    }

    let html = '';
    const sortedComponents = [...report.components].sort((a, b) => a.order - b.order);
    
    for (let i = 0; i < sortedComponents.length; i++) {
        const comp = sortedComponents[i];
        const nextComp = sortedComponents[i + 1];
        const isHalf = comp.width === 'half';
        const nextIsHalf = nextComp?.width === 'half';
        
        const value1 = draftData[comp.order] !== undefined ? draftData[comp.order] : comp.value;

        if (isHalf && nextIsHalf) {
            const value2 = draftData[nextComp.order] !== undefined ? draftData[nextComp.order] : nextComp.value;
            html += `<div class="grid grid-cols-1 md:grid-cols-2 gap-4"><div>${renderSinglePreviewComponent(comp, true, value1)}</div><div>${renderSinglePreviewComponent(nextComp, true, value2)}</div></div>`;
            i++;
        } else {
            html += renderSinglePreviewComponent(comp, true, value1);
        }
    }

    // APLICANDO A CORREÇÃO FIXA DO RODAPÉ (com nova estrutura)
    elements.fillReportModal.querySelector('.modal').classList.add('modal-fill-scroll');


    elements.fillReportForm.innerHTML = `
        <div id="fill-report-content-scroll" class="space-y-6">
           ${html}
        </div>
        <div class="fixed-modal-footer">
            <div class="flex justify-end gap-3 pt-4 border-t dark:border-gray-700">
                <button type="button" id="save-draft-btn-modal" class="btn btn-warning"><i data-lucide="save" class="w-5 h-5"></i> Salvar Rascunho</button>
                <button type="button" id="cancel-fill-report-btn-modal" class="btn btn-dark">Cancelar</button>
                <button type="submit" class="btn btn-primary"><i data-lucide="send" class="w-5 h-5"></i> Enviar Relatório</button>
            </div>
        </div>
    `;

    document.getElementById('save-draft-btn-modal').addEventListener('click', handleSaveDraft);
    document.getElementById('cancel-fill-report-btn-modal').addEventListener('click', closeFillReportModal);
    
    try { lucide.createIcons(); } catch(e) {}
    elements.fillReportModal.classList.add('active');
}

function viewFilledReport(reportId) {
    const report = state.filledReports.find(r => r.id === reportId);
    if (!report) { showToast("Relatório preenchido não encontrado.", "error"); return; }
    
    elements.viewReportModalTitle.textContent = report.templateTitle;
    elements.viewReportFilledBy.textContent = report.filledByName || report.filledByEmail;
    elements.viewReportFilledAt.textContent = new Date(report.filledAt).toLocaleString();
    
    let html = '';
    (report.data || []).forEach(item => {
        let valueDisplay;
        if (Array.isArray(item.value)) {
            if (item.componentType === 'table') {
                const columns = Object.keys(item.value[0] || {});
                if (columns.length > 0) {
                    valueDisplay = `<div class="overflow-x-auto"><table class="min-w-full text-sm border border-gray-200 dark:border-gray-700 rounded-lg"><thead><tr class="bg-gray-100 dark:bg-gray-700 text-left">${columns.map(col => `<th class="px-4 py-2">${col}</th>`).join('')}</tr></thead><tbody>${item.value.map(row => `<tr>${columns.map(col => `<td class="px-4 py-2">${row[col] || '-'}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
                } else {
                    valueDisplay = '<span class="text-gray-500 italic">Tabela vazia.</span>';
                }
            } else {
                valueDisplay = `<ul class="list-disc pl-5 space-y-1 text-gray-700 dark:text-gray-300">${item.value.map(val => `<li>${val}</li>`).join('')}</ul>`;
            }
        } else if (!item.value) {
            valueDisplay = '<span class="text-gray-500 italic">Não preenchido.</span>';
        } else if (item.componentType === 'image' || item.componentType === 'file') {
            valueDisplay = `<p class="font-medium text-blue-600 dark:text-blue-400">Anexo: ${item.value} <i data-lucide="${item.componentType === 'image' ? 'image' : 'paperclip'}" class="w-4 h-4 inline ml-1"></i></p>`;
        } else {
            valueDisplay = `<p class="font-medium text-gray-900 dark:text-white whitespace-pre-wrap">${item.value}</p>`;
        }

        html += `<div class="p-3 border-b dark:border-gray-700 last:border-b-0">
            <h4 class="text-sm font-bold dark:text-gray-200">${item.componentLabel} (${getComponentTypeInfo(item.componentType).name})</h4>
            <div class="mt-1">${valueDisplay}</div>
        </div>`;
    });
    
    elements.viewReportContent.innerHTML = html;
    
    // CORREÇÃO CRÍTICA: Adiciona os botões de download e impressão ao footer do modal.
    const footerContainer = elements.viewReportModal.querySelector('.flex.justify-end.gap-3.pt-4.mt-6.border-t');
    
    footerContainer.querySelectorAll('.btn:not(#close-view-report-btn)').forEach(btn => btn.remove());

    const buttonsHtml = `
        <button type="button" class="btn btn-secondary print-hide" onclick="window.print()"><i data-lucide="printer" class="w-5 h-5"></i> Imprimir</button>
        <button type="button" class="btn btn-success print-hide" onclick="showToast('Download PDF (simulado)', 'info')"><i data-lucide="file-text" class="w-5 h-5"></i> PDF</button>
        <button type="button" class="btn btn-success print-hide" onclick="showToast('Download DOCX (simulado)', 'info')"><i data-lucide="file-text" class="w-5 h-5"></i> DOCX</button>
    `;
    
    elements.closeViewReportBtn.insertAdjacentHTML('beforebegin', buttonsHtml);
    try { lucide.createIcons(); } catch(e) {}

    elements.viewReportModal.classList.add('active');
}

function deleteDraft(id) {
    showConfirmModal("Excluir Rascunho", "Tem certeza que deseja excluir este rascunho?", () => {
        if (typeof deleteDraftFromFirestore === 'function' && state.currentUser.id !== 'local_admin') {
            deleteDraftFromFirestore(id)
               .then(() => showToast('Rascunho excluído!'))
               .catch(() => showToast('Erro ao excluir rascunho.', 'error'));
        } else {
            state.draftReports = state.draftReports.filter(d => d.id !== id); 
            state.notifications.draftReports = state.notifications.draftReports - 1;
            updateNotificationBell(); 
            renderDraftReports(); 
            showToast('Rascunho excluído (Local).');
        }
    }, 'btn-danger');
}
function renderSavedReports() {
    if (!elements.savedReportsList) { return; }
    if (state.savedReports.length === 0) {
        elements.savedReportsList.innerHTML = `<div class="col-span-full text-center py-10 text-gray-500 dark:text-gray-400"><i data-lucide="folder-open" class="w-12 h-12 mx-auto mb-3 opacity-50"></i><p>Nenhum modelo de relatório criado.</p></div>`;
        try { lucide.createIcons(); } catch(e) {}
        return;
    }
    elements.savedReportsList.innerHTML = state.savedReports.map(report => {
        const isDisabled = report.status === 'paused'; const statusIcon = isDisabled ? 'play' : 'pause'; const statusTitle = isDisabled ? 'Ativar' : 'Pausar'; const statusBtnClass = isDisabled ? 'btn-success' : 'btn-warning';
        const canEdit = state.currentUser?.permissions?.canCreateReports;
        const canFill = state.currentUser?.permissions?.canFillReports;

        return `<div class="card p-4 ${isDisabled ? 'report-card-disabled' : ''}"><div class="flex items-start justify-between"><div><h3 class="font-bold text-gray-800 dark:text-white">${report.title}</h3><p class="text-sm text-gray-500 mt-1">${report.components?.length || 0} componentes • Criado por ${report.createdBy || 'Sistema'}</p></div><i data-lucide="file-text" class="w-5 h-5 text-blue-500 shrink-0"></i></div><div class="flex gap-2 mt-4">${canFill ? `<button class="btn btn-sm btn-primary flex-1" onclick="fillReport('${report.id}')"><i data-lucide="edit" class="w-4 h-4"></i> Preencher</button>` : ''}${canEdit ? `<button class="btn btn-sm btn-secondary" onclick="loadReportForEditing('${report.id}')" title="Editar Modelo"><i data-lucide="settings" class="w-4 h-4"></i></button><button class="btn btn-sm ${statusBtnClass}" onclick="toggleReportStatus('${report.id}')" title="${statusTitle}"><i data-lucide="${statusIcon}" class="w-4 h-4"></i></button><button class="btn btn-sm btn-danger" onclick="deleteReport('${report.id}')" title="Excluir Modelo"><i data-lucide="trash-2" class="w-4 h-4"></i></button>` : ''}</div></div>`;
    }).join('');
    try { lucide.createIcons(); } catch(e) {}
}
function renderFilledReports() {
     if (!elements.filledReportsList) return;
     if (state.filledReports.length === 0) {
         elements.filledReportsList.innerHTML = `<div class="text-center py-10 text-gray-500 dark:text-gray-400"><i data-lucide="bar-chart-3" class="w-12 h-12 mx-auto mb-3 opacity-50"></i><p>Nenhum relatório preenchido.</p></div>`;
         try{lucide.createIcons();}catch(e){}
     } else {
         elements.filledReportsList.innerHTML = state.filledReports.sort((a,b) => new Date(b.filledAt) - new Date(a.filledAt)).map(report => `
             <div class="card p-4">
                 <h4 class="font-semibold dark:text-white">${report.templateTitle}</h4>
                 <p class="text-xs text-gray-500">Preenchido por ${report.filledByName || report.filledByEmail} em ${new Date(report.filledAt).toLocaleString()}</p>
                 <div class="mt-3 flex gap-2">
                     <button class="btn btn-sm btn-secondary" onclick="viewFilledReport('${report.id}')"><i data-lucide="eye" class="w-4 h-4"></i> Visualizar</button>
                 </div>
             </div>`).join('');
         try { lucide.createIcons(); } catch(e) {}
     }
}
function renderDraftReports() {
    if (!elements.draftsList) return;
    if (state.draftReports.length === 0) { elements.draftsListContainer.classList.add('hidden'); return; }
    elements.draftsListContainer.classList.remove('hidden');
    elements.draftsList.innerHTML = state.draftReports.sort((a,b) => new Date(b.lastUpdatedAt) - new Date(a.lastUpdatedAt)).map(draft => `
        <div class="card p-4 bg-yellow-50 dark:bg-yellow-900/20 border-yellow-200 dark:border-yellow-800">
            <h4 class="font-semibold text-yellow-800 dark:text-yellow-200">${draft.templateTitle} (Rascunho)</h4>
            <p class="text-xs text-yellow-700 dark:text-yellow-400">Última edição por ${draft.createdByName || draft.createdByEmail} em ${new Date(draft.lastUpdatedAt).toLocaleString()}</p>
            <div class="mt-3 flex gap-2">
                <button class="btn btn-sm btn-primary" onclick="fillReport('${draft.templateId}', '${draft.id}')"><i data-lucide="edit" class="w-4 h-4"></i> Continuar</button>
                <button class="btn btn-sm btn-danger" onclick="deleteDraft('${draft.id}')"><i data-lucide="trash-2" class="w-4 h-4"></i> Excluir</button>
            </div>
        </div>`).join('');
    try { lucide.createIcons(); } catch(e) {}
}
function renderUsersTable() {
     if (!elements.usersTableBody) return; 
     if (state.currentUser?.profile === 'admin') { renderPendingUsers(); }
     
     const usersToRender = state.users.filter(u => u.id !== 'local_admin' || u.id === state.currentUser?.id);

     if (usersToRender.length === 0) { elements.usersTableBody.innerHTML = `<tr><td colspan="4" class="text-center py-10 text-gray-500"><i data-lucide="users" class="w-12 h-12 mx-auto mb-3 opacity-50"></i><p>Nenhum usuário.</p></td></tr>`; try{lucide.createIcons();}catch(e){} return; }
     elements.usersTableBody.innerHTML = usersToRender.map(user => {
          const isAuthorized = user.autorizado === undefined || user.autorizado === true;
          let statusBadge = '';
          if (!isAuthorized) { statusBadge = `<span class="px-2 py-1 text-xs rounded-full bg-yellow-100 text-yellow-800 dark:bg-yellow-900 dark:text-yellow-200 ml-2">Pendente</span>`; }
          
          const isProtectedAdmin = user.profile === 'admin' && (user.id === 'local_admin' || user.id === state.currentUser?.id);
          const isCurrentUser = user.id === state.currentUser?.id;

          return `<tr class="border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-800"><td class="py-3 px-4"><div class="flex items-center gap-3"><img src="https://ui-avatars.com/api/?name=${encodeURIComponent(user.name||'U')}&background=random&color=fff" alt="${user.name}" class="w-8 h-8 rounded-full"><div><p class="font-medium dark:text-white">${user.name||'S/ Nome'}${isCurrentUser ? ' (Você)' : ''}</p></div></div></td><td class="py-3 px-4 text-gray-600 dark:text-gray-400">${user.email}</td><td class="py-3 px-4"><span class="px-2 py-1 text-xs rounded-full ${getProfileBadgeClass(user.profile)}">${getProfileName(user.profile)}</span>${statusBadge}</td><td class="py-3 px-4 text-center"><div class="flex justify-center gap-2">${!isProtectedAdmin ? `<button class="btn btn-sm btn-secondary" onclick="editUser('${user.id}')"><i data-lucide="edit"></i></button><button class="btn btn-sm btn-danger" onclick="deleteUser('${user.id}')"><i data-lucide="trash-2" class="w-4 h-4"></i></button>` : `<span class="text-xs text-gray-400 italic">Protegido</span>`}</div></td></tr>`;
     }).join('');
     try { lucide.createIcons(); } catch(e) {}
}
function renderPendingUsers() {
     if (!elements.usersTableBody) return;
     if (!state.currentUser || state.currentUser.profile !== 'admin') { elements.pendingUsersContainer.classList.add('hidden'); return; }

     if (state.pendingUsers.length === 0) { elements.pendingUsersContainer.classList.add('hidden'); return; }
     elements.pendingUsersContainer.classList.remove('hidden');
     elements.pendingUsersTableBody.innerHTML = state.pendingUsers.map(user => `
          <tr class="border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-800">
               <td class="py-3 px-4">${user.name}</td>
               <td class="py-3 px-4 text-gray-600 dark:text-gray-400">${user.email}</td>
               <td class="py-3 px-4 text-center">
                    <div class="flex justify-center gap-2">
                         <button class="btn btn-sm btn-success" title="Aprovar" onclick="approveUser('${user.id}')"><i data-lucide="check"></i></button>
                         <button class="btn btn-sm btn-danger" title="Negar" onclick="denyUser('${user.id}')"><i data-lucide="x"></i></button>
                    </div>
               </td>
          </tr>`).join('');
     try { lucide.createIcons(); } catch(e) {}
}
function showAddUserModal() {
    elements.userModalTitle.textContent = 'Adicionar Novo Usuário';
    elements.userForm.reset();
    elements.userForm.querySelector('#user-id').value = '';
    
    document.getElementById('password-field-container').classList.remove('hidden');
    elements.userPasswordInput.setAttribute('required', 'required');

    elements.userForm.querySelector('#user-profile').value = 'user';
    elements.userForm.querySelectorAll('input[type="checkbox"][data-permission]').forEach(check => {
        const perm = check.dataset.permission;
        check.checked = userPermissions[perm] || false; 
    });

    elements.userModal.classList.add('active');
}
function editUser(id) {
    const user = state.users.find(u => u.id == id);
    if (!user) { showToast('Usuário não encontrado.', 'error'); return; }
    
    elements.userModalTitle.textContent = `Editar Usuário: ${user.name}`;
    elements.userForm.querySelector('#user-id').value = user.id;
    elements.userForm.querySelector('#user-name').value = user.name;
    elements.userForm.querySelector('#user-email').value = user.email;
    elements.userForm.querySelector('#user-profile').value = user.profile;
    
    document.getElementById('password-field-container').classList.add('hidden');
    elements.userPasswordInput.removeAttribute('required');

    elements.userForm.querySelectorAll('input[type="checkbox"][data-permission]').forEach(check => {
        const perm = check.dataset.permission;
        check.checked = user.permissions[perm] || false;
    });

    elements.userModal.classList.add('active');
}
async function updateUserAuthorization(uid, authorized) {
    if (!isFirebaseReady) { showToast("Firebase não pronto. Operação local não suportada.", "error"); return; }
    
    const user = state.pendingUsers.find(u => u.id === uid);
    if (!user) { showToast("Usuário pendente não encontrado.", "error"); return; }

    const userRef = doc(db, USERS_COLLECTION_PATH, uid);
    
    if (authorized) {
        await updateDoc(userRef, { 
            autorizado: true,
            permissions: user.permissions || userPermissions
        }).then(() => {
            showToast(`Usuário ${user.name} aprovado com sucesso!`, 'success');
            logActivity("user-approved", { targetUid: uid, targetEmail: user.email });
        }).catch((e) => {
            showToast("Erro ao aprovar usuário.", 'error');
            console.error("Approval error:", e);
        });
    } else {
        showConfirmModal("Negar Acesso", `Tem certeza que deseja negar e EXCLUIR o registro de ${user.name}?`, async () => {
            await deleteDoc(doc(db, USERS_COLLECTION_PATH, uid)).then(() => {
                showToast(`Solicitação de ${user.name} negada e registro excluído!`, 'success');
                logActivity("user-denied", { targetUid: uid, targetEmail: user.email });
            }).catch((e) => {
                showToast("Erro ao negar/excluir usuário.", 'error');
                console.error("[Firestore] Denial/Delete error:", e);
            });
        }, 'btn-danger');
    }
}

// CORREÇÃO: Função de ditado para Textarea
function handleDictationClick(buttonElement) {
    const targetId = buttonElement.dataset.target;
    // O targetInput é o campo de entrada exato (textarea ou input)
    const targetInput = document.getElementById(targetId);
    
    if (!targetInput) {
        showToast('Campo de destino não encontrado.', 'error');
        return;
    }

    if (!state.speechRecognition) {
        initSpeechRecognition();
        if (!state.speechRecognition) return;
    }

    // 1. Lógica de Início/Parada (se o mesmo botão for clicado)
    if (buttonElement.classList.contains('listening')) {
        state.speechRecognition.stop();
        buttonElement.classList.remove('listening');
        showToast('Ditado parado.', 'info');
        return;
    }

    // 2. Parar instância anterior (se outro botão foi clicado)
    if (state.speechTarget) {
        state.speechRecognition.stop();
        const prevButton = state.speechTarget.parentElement.querySelector('.dictation-button.listening');
        if (prevButton) {
            prevButton.classList.remove('listening');
        }
    }

    // 3. Configurar e iniciar nova gravação
    state.speechTarget = targetInput;
    buttonElement.classList.add('listening');

    state.speechRecognition.onresult = (event) => {
        let finalTranscript = '';

        for (let i = 0; i < event.results.length; ++i) {
            if (event.results[i].isFinal) {
                finalTranscript += event.results[i][0].transcript;
            }
        }
        
        // CORREÇÃO CRÍTICA: Adicionar o novo texto final ao valor existente do campo correto.
        targetInput.focus();
        
        // Adiciona espaço ANTES se o campo já tiver conteúdo e não houver pontuação de espaço
        const currentValue = targetInput.value;
        if (currentValue.length > 0 && finalTranscript.trim().length > 0 && 
            !currentValue.endsWith(' ') && 
            !currentValue.match(/[\.\,\?\!]$/)) {
            targetInput.value += ' ';
        }
        targetInput.value += finalTranscript;
    };

    state.speechRecognition.onerror = (event) => {
        showToast(`Erro no ditado: ${event.error}`, 'error');
        buttonElement.classList.remove('listening');
        state.speechTarget = null;
    };

    state.speechRecognition.onend = () => {
        buttonElement.classList.remove('listening');
        state.speechTarget = null;
        showToast('Ditado finalizado.', 'info');
    };

    state.speechRecognition.start();
    showToast('Comece a falar agora...', 'info');
}

// CORREÇÃO: Inicialização da API de Reconhecimento de Fala
function setupSpeechRecognition() { 
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

    if (SpeechRecognition) {
        state.speechRecognition = new SpeechRecognition();
        state.speechRecognition.continuous = true; // Para capturar frases longas
        state.speechRecognition.interimResults = false; // Apenas resultados finais
        state.speechRecognition.lang = 'pt-BR'; // Configurar para português

        console.log("Speech Recognition setup complete.");
    } else {
        console.log("Speech Recognition setup skipped (API not supported).");
    }
}
function setupFillReportEventListeners() {
    console.log("Fill Report event listeners initialized.");
}


// ----------------------------------------------------------------------
// --- CÓDIGO DE INICIALIZAÇÃO (executado ao carregar o DOM) ------------
// ----------------------------------------------------------------------

document.addEventListener('DOMContentLoaded', () => {
    // 1. Inicialização UI básica
    if (state.darkMode) { document.documentElement.classList.add('dark'); if(elements.themeToggle) elements.themeToggle.classList.add('active'); }
    else { document.documentElement.classList.remove('dark'); if(elements.themeToggle) elements.themeToggle.classList.add('active'); }
    
    const loginForm = document.getElementById('login-form');
    
    // --- Criação do botão Google ---
    const googleBtn = document.createElement('button');
    googleBtn.id = 'login-google-btn'; googleBtn.type = 'button';
    googleBtn.className = 'btn w-full mt-4 bg-white text-gray-700 hover:bg-gray-100 dark:bg-gray-800 dark:text-gray-200 dark:hover:bg-gray-700 border border-gray-300 dark:border-gray-600 shadow-md';
    googleBtn.innerHTML = '<i data-lucide="mail" class="w-5 h-5"></i> Entrar com Google';
    document.querySelector('#login-screen .bg-white .mt-6')?.appendChild(googleBtn);
    
    // --- Criação e Listener do Toggle de Senha de Login ---
    if (elements.loginPasswordInput) {
        const passwordDiv = elements.loginPasswordInput.closest('div');
        if (passwordDiv && !document.getElementById('toggle-login-password-visibility')) {
            passwordDiv.classList.add('relative'); 
            const toggleBtn = document.createElement('button');
            toggleBtn.type = 'button'; toggleBtn.className = 'absolute right-3 top-1/2 transform -translate-y-1/2 text-gray-500 hover:text-gray-700';
            toggleBtn.id = 'toggle-login-password-visibility'; toggleBtn.title = 'Mostrar/Esconder senha';
            toggleBtn.innerHTML = '<i data-lucide="eye" class="w-5 h-5"></i>';
            passwordDiv.appendChild(toggleBtn);
            elements.toggleLoginPasswordVisibility = toggleBtn;

            toggleBtn.addEventListener('click', (e) => { e.preventDefault(); togglePasswordVisibility(elements.loginPasswordInput, e.currentTarget); });
        }
    }

    // --- Criação e Listener do Toggle de Senha do Modal ---
    elements.togglePasswordVisibility = document.getElementById('toggle-password-visibility');
    if (elements.togglePasswordVisibility && elements.userPasswordInput) {
        elements.togglePasswordVisibility.addEventListener('click', (e) => { e.preventDefault(); togglePasswordVisibility(elements.userPasswordInput, e.currentTarget); });
    }
    
    // 3. Setup de Event Listeners
    elements.loginForm?.addEventListener('submit', handleLogin);
    elements.logoutBtn?.addEventListener('click', handleLogout);
    elements.themeToggle?.addEventListener('click', toggleTheme);
    elements.menuToggle?.addEventListener('click', toggleSidebar);
    elements.saveReportBtn?.addEventListener('click', saveReport);
    elements.cancelEditBtn?.addEventListener('click', () => cancelReportEdit(true));
    elements.addUserBtn?.addEventListener('click', showAddUserModal);
    elements.cancelUserBtn?.addEventListener('click', closeUserModal);
    elements.userForm?.addEventListener('submit', handleUserFormSubmit);
    elements.reportTitle?.addEventListener('input', updatePreviewTitle);
    elements.componentForm?.addEventListener('submit', handleComponentFormSubmit);
    elements.cancelComponentBtn?.addEventListener('click', closeComponentModal);
    elements.fillReportForm?.addEventListener('submit', handleFillReportSubmit);
    elements.confirmBtn?.addEventListener('click', () => { if (typeof confirmCallback === 'function') { try { confirmCallback(); } catch (e) { console.error("Error in confirm callback:", e); } } closeConfirmModal(); });
    elements.confirmCancelBtn?.addEventListener('click', closeConfirmModal);
    elements.confirmModalEl?.addEventListener('click', (event) => { if (event.target === elements.confirmModalEl) closeConfirmModal(); });
    elements.notificationBellButton?.addEventListener('click', toggleNotificationDropdown);
    elements.notificationList?.addEventListener('click', handleNotificationClick);
    elements.closeViewReportBtn?.addEventListener('click', closeViewReportModal);
    elements.viewReportModal?.addEventListener('click', (event) => { if (event.target === elements.viewReportModal) closeViewReportModal(); });
    
    document.querySelectorAll('#main-nav .nav-item').forEach(item => { item.addEventListener('click', (e) => { e.preventDefault(); switchPage(item.dataset.page); }); });

    // Setup de componentes e listeners dinâmicos
    setupBuilderEventListeners();
    setupFillReportEventListeners();
    
    // 4. Inicialização de serviços
    initFirebase();
    initSortable(); 
    setupSpeechRecognition();

    setTimeout(() => { try { lucide.createIcons(); } catch(e) {} }, 100);

    console.log("[Super App] Inicialização do módulo principal concluída.");
});

// ----------------------------------------------------------------------
// --- EXPOSIÇÃO DE FUNÇÕES GLOBAIS (para onclick no HTML) --------------
// ----------------------------------------------------------------------

window.signInWithGoogle = signInWithGoogle;
window.switchPage = switchPage;
window.editUser = editUser;
window.deleteUser = deleteUser;
window.approveUser = (uid) => { showConfirmModal("Aprovar Usuário", "Tem certeza que deseja aprovar este usuário?", () => { updateUserAuthorization(uid, true); }, 'btn-success'); };
window.denyUser = (uid) => { showConfirmModal("Negar Usuário", "Tem certeza que deseja negar este usuário?", () => { updateUserAuthorization(uid, false); }, 'btn-danger'); };
window.fillReport = fillReport;
window.loadReportForEditing = loadReportForEditing;
window.toggleReportStatus = toggleReportStatus;
window.deleteReport = deleteReport;
window.deleteDraft = deleteDraft;
window.viewFilledReport = viewFilledReport;
window.closeViewReportModal = closeViewReportModal;
window.showToast = showToast;
window.handleTableAddRow = handleTableAddRow;
window.handleTableRemoveRow = handleTableRemoveRow;
window.handleFileUpload = handleFileUpload;

window.handleDictationClick = handleDictationClick;
