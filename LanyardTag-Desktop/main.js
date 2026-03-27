const { app, BrowserWindow, shell, ipcMain } = require('electron');
const path = require('path');

function createWindow() {
  const win = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1024,
    minHeight: 700,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
    title: '줄명찰 시스템',
    show: false,
    backgroundColor: '#0d1117',
  });

  win.loadFile('index.html');
  win.once('ready-to-show', () => win.show());

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

// 인쇄 HTML을 별도 창으로 열어 네이티브 인쇄 다이얼로그 사용 (Windows 최적화)
ipcMain.handle('print-html', async (event, html) => {
  return new Promise((resolve) => {
    const printWin = new BrowserWindow({
      width: 900,
      height: 1200,
      show: false,
      webPreferences: { nodeIntegration: false, contextIsolation: true },
    });

    printWin.loadURL('about:blank');

    printWin.webContents.once('did-finish-load', async () => {
      await printWin.webContents.executeJavaScript(
        `document.open('text/html','replace'); document.write(${JSON.stringify(html)}); document.close();`
      );

      // 이미지 로드 대기
      setTimeout(() => {
        printWin.webContents.print(
          { silent: false, printBackground: true, color: true },
          (success, reason) => {
            printWin.close();
            resolve({ success, reason: reason || '' });
          }
        );
      }, 1200);
    });
  });
});

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  app.quit();
});
