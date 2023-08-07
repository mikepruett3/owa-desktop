// main.js

// https://www.electronforge.io/config/makers/squirrel.windows
if (require('electron-squirrel-startup')) return;

const { app, shell, BrowserWindow, Menu, MenuItem, Tray, nativeImage, dialog } = require('electron');
const { getHA, setHA } = require('./settings.js');

// Disable Hardware Acceleration
// https://www.electronjs.org/docs/latest/tutorial/offscreen-rendering
if (!getHA()) {
    app.disableHardwareAcceleration()
}

createWindow = () => {
    const window = new BrowserWindow({
        width: 1280,
        height: 720,
        title: 'Outlook Desktop',
        icon: __dirname + '/images/Outlook.ico',
        autoHideMenuBar: true,
        webPreferences: {
            spellcheck: true,
            webSecurity: true,
            contextIsolation: false,
            webviewTag: true,
            nodeIntegration: true,
            nativeWindowOpen: true
        }
    });

    window.loadURL(`https://outlook.office.com/`);

    // Open links with External Browser
    // https://stackoverflow.com/a/67409223
    window.webContents.setWindowOpenHandler(({ url }) => {
        shell.openExternal(url);
        return { action: 'deny' };
    });

    window.webContents.on('context-menu', (event, params) => {
        const menu = new Menu()

        // Add each spelling suggestion
        for (const suggestion of params.dictionarySuggestions) {
            menu.append(
                new MenuItem({
                    label: suggestion,
                    click: () => window.webContents.replaceMisspelling(suggestion)
                })
            )
        }

        // Allow users to add the misspelled word to the dictionary
        if (params.misspelledWord) {
            menu.append(
                new MenuItem({
                    label: 'Add to dictionary',
                    click: () => window.webContents.session.addWordToSpellCheckerDictionary(params.misspelledWord)
                })
            )
        }

        menu.popup()
    })

    const icon = nativeImage.createFromPath(__dirname + '/images/Outlook.ico')
    const tray = new Tray(icon)

    const contextMenu = Menu.buildFromTemplate([
        {
            label: 'Default Mailto Client',
            type: 'checkbox',
            checked: app.isDefaultProtocolClient('mailto'),
            click() {
                if (app.isDefaultProtocolClient('mailto')) {
                    app.removeAsDefaultProtocolClient('mailto')
                } else {
                    app.setAsDefaultProtocolClient('mailto')
                }
            }
        },
        {
            label: 'Hardware Acceleration',
            type: 'checkbox',
            checked: getHA(),
            click({ checked }) {
                setHA(checked)
                dialog.showMessageBox(
                    null,
                    {
                        type: 'info',
                        title: 'info',
                        message: 'Exiting Applicatiom, as Hardware Acceleration setting has been changed...'

                    })
                    .then(result => {
                      if (result.response === 0) {
                        app.relaunch();
                        app.exit()
                      }
                    }
                );
            }
        },
        {
            label: 'Clear Cache',
            click: () => {
                session.defaultSession.clearStorageData()
                app.relaunch();
                app.exit();
            }
        },
        {
            label: 'Reload',
            click: () => win.reload()
        },
        {
            label: 'Quit',
            type: 'normal',
            role: 'quit'
        }
    ])

    tray.setToolTip('Outlook Desktop')
    tray.setTitle('Outlook Desktop')
    tray.setContextMenu(contextMenu)
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