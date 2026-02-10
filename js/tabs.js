/**
 * SumCheck - Tab 切換控制器
 * 管理「載入檔案」與「貼上內容」兩個頁面的切換
 */

(function () {
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPages = document.querySelectorAll('.tab-page');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.dataset.tab;

            // 切換按鈕狀態
            tabBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            // 切換頁面
            tabPages.forEach(page => {
                page.classList.toggle('hidden', page.id !== targetId);
            });

            // 首次切換到貼上頁時初始化 PasteApp
            if (targetId === 'paste-page' && typeof PasteApp !== 'undefined' && !PasteApp._initialized) {
                PasteApp.init();
                PasteApp._initialized = true;
            }
        });
    });
})();
