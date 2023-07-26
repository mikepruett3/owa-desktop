// main.js

// https://www.electronforge.io/config/makers/squirrel.windows
//if (require('electron-squirrel-startup')) return;

const { app, shell, BrowserWindow, Menu, MenuItem } = require('electron');

// Disable Hardware Acceleration
// https://www.electronjs.org/docs/latest/tutorial/offscreen-rendering
app.disableHardwareAcceleration()

createWindow = () => {
    const window = new BrowserWindow({
        width: 1280,
        height: 720,
        title: 'Outlook Desktop',
        icon: __dirname + '/images/Outlook.ico',
        autoHideMenuBar: true,
        webPreferences: {
            webSecurity: true,
            contextIsolation: false,
            webviewTag: true,
            nodeIntegration: true,
            nativeWindowOpen: true
        }
    });

    // Open links with External Browser
    // https://stackoverflow.com/a/67409223
    window.webContents.setWindowOpenHandler(({ url }) => {
        shell.openExternal(url);
        return { action: 'deny' };
    });

    window.loadURL(`https://outlook.office.com/`);
};

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});