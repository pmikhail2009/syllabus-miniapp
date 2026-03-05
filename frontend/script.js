const tg = window.Telegram.WebApp;
tg.expand();

let appState = { grade: null, subject: null };
let syllabusData = {};

// Загрузка данных из JSON
async function loadData() {
    try {
        const res = await fetch('data.json');
        syllabusData = await res.json();
    } catch (e) {
        console.error('Ошибка загрузки data.json:', e);
        tg.showPopup({ title: 'Ошибка', message: 'Не удалось загрузить данные', buttons: [{ type: 'ok' }] });
    }
}

function showScreen(name) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    document.getElementById(name + '-screen').classList.add('active');
}

function showLoading(show) {
    document.getElementById('loading').classList.toggle('hidden', !show);
}

// Выбор класса
document.querySelectorAll('.grade-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        appState.grade = btn.dataset.grade;
        const subjects = Object.keys(syllabusData[appState.grade] || {}).sort();
        
        const container = document.getElementById('subjects-list');
        container.innerHTML = '';
        subjects.forEach(subj => {
            const item = document.createElement('div');
            item.className = 'subject-item';
            item.innerHTML = `<span class="subject-name">${subj}</span><span>›</span>`;
            item.onclick = () => selectSubject(subj);
            container.appendChild(item);
        });
        
        document.getElementById('grade-title').textContent = `${appState.grade} класс — Предметы`;
        showScreen('subject');
        if (tg.HapticFeedback) tg.HapticFeedback.impactOccurred('light');
    });
});

// Выбор предмета
function selectSubject(subject) {
    appState.subject = subject;
    const teachers = syllabusData[appState.grade][subject];
    
    if (teachers.length === 1) {
        showResult(subject, appState.grade, teachers[0]);
    } else {
        const container = document.getElementById('teachers-list');
        container.innerHTML = '';
        teachers.forEach(t => {
            const item = document.createElement('div');
            item.className = 'teacher-item';
            item.innerHTML = `<span class="teacher-name">${t.teacher}</span><span>›</span>`;
            item.onclick = () => showResult(subject, appState.grade, t);
            container.appendChild(item);
        });
        document.getElementById('subject-title').textContent = `"${subject}" — выберите преподавателя:`;
        showScreen('teacher');
    }
    if (tg.HapticFeedback) tg.HapticFeedback.impactOccurred('light');
}

// Показать результат
function showResult(subject, grade, teacherData) {
    document.getElementById('result-subject').textContent = subject;
    document.getElementById('result-grade').textContent = grade;
    document.getElementById('result-teacher').textContent = teacherData.teacher;
    document.getElementById('result-link').href = teacherData.url;
    
    tg.MainButton.setText('📄 Открыть силлабус');
    tg.MainButton.onClick(() => tg.openLink(teacherData.url));
    tg.MainButton.show();
    
    showScreen('result');
    if (tg.HapticFeedback) tg.HapticFeedback.notificationOccurred('success');
}

// Навигация
function goBack() {
    if (document.getElementById('teacher-screen').classList.contains('active')) {
        showScreen('subject');
    } else {
        appState = { grade: null, subject: null };
        tg.MainButton.hide();
        showScreen('grade');
    }
}

function resetApp() {
    appState = { grade: null, subject: null };
    tg.MainButton.hide();
    showScreen('grade');
}

// Инициализация
document.addEventListener('DOMContentLoaded', async () => {
    await loadData();
    tg.ready();
});