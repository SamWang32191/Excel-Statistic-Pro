/**
 * 主題切換邏輯
 * 支援：淺色、深色、跟隨系統
 */

export function initTheme() {
    const themeToggle = document.getElementById('themeToggle');
    if (!themeToggle) return;

    const themeButtons = themeToggle.querySelectorAll('.btn-theme');
    const savedTheme = localStorage.getItem('theme') || 'system';

    // 套用初始主題
    applyTheme(savedTheme);

    // 監聽按鈕點擊
    themeButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const theme = btn.getAttribute('data-theme');
            applyTheme(theme);
            localStorage.setItem('theme', theme);
        });
    });

    // 監聽系統主題變更
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
        if (localStorage.getItem('theme') === 'system' || !localStorage.getItem('theme')) {
            applyTheme('system');
        }
    });
}

function applyTheme(theme) {
    const html = document.documentElement;
    const themeButtons = document.querySelectorAll('.btn-theme');

    // 移除所有按鈕的 active 狀態
    themeButtons.forEach(btn => btn.classList.remove('active'));

    // 獲取當前應顯示的主題 (若為 system 則判斷電腦設定)
    let actualTheme = theme;
    if (theme === 'system') {
        actualTheme = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    }

    // 更新 HTML 屬性
    if (actualTheme === 'dark') {
        html.removeAttribute('data-theme'); // 預設為深色，不需屬性
    } else {
        html.setAttribute('data-theme', 'light');
    }

    // 更新按鈕 active 狀態
    const activeBtn = document.querySelector(`.btn-theme[data-theme="${theme}"]`);
    if (activeBtn) activeBtn.classList.add('active');

    // 記錄目前的 mode 到 body 類別 (可選)
    document.body.classList.remove('theme-dark', 'theme-light');
    document.body.classList.add(`theme-${actualTheme}`);
}
