# Outlook Desktop

Outlook Desktop is a simple  Outlook Web (Office 365) Desktop application, built using [ElectronJS](https://www.electronjs.org).

## Installation

Dowload the lates [release](https://github.com/mikepruett3/owa-desktop/releases) for Windows. Unfortunatley I am not able to provide packages for Linux or MacOS at this time.

For Windows... a standard Exectuable is provided, as well as a NuGet package.

## Launching

To run, just launch the executable via the Desktop Shortcut, or the Executable directly.

Hardware Acceleration is disabled at launch

## Building

To build locally, clone the repository and install the dependencies.

```powershell
git clone https://github.com/mikepruett3/owa-desktop.git
cd owa-desktop
npm install
```

To run the application locally.

```powershell
npm run test
```

To build the application installer.

```powershell
npm run make
```

## Dependencies

- electron-squirrel-startup
- publisher-github
- electron-forge
- electron
