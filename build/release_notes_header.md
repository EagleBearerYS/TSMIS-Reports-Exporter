Portable Windows desktop app that bulk-exports Caltrans TSMIS reports for every
California state route — sign in once, tick the reports you want, done (see
`Start Here.txt` inside the zip).

## Pick your download

| Download | Best for | Browser it uses |
|---|---|---|
| `…-win64.zip` | Most users — smallest download | The Google Chrome / Microsoft Edge already installed on the PC |
| `…-win64-with-browser.zip` | Managed PCs where Edge sign-in misbehaves and Chrome isn't installed | Ships its own **Built-in Chromium** and uses it by default (Settings ▸ Export browser can switch to Google Chrome) |
| `…-batch-source.zip` | Developers / console fallback | Requires Python 3.11 — `1. setup (one time).bat` installs the libraries **and downloads Chromium** |

Both app zips: unzip into a writable folder with a **short path** (`C:\TSMIS` is
ideal — Windows and Excel limit how long a file's full path can be) and
double-click `TSMIS Exporter.exe`.

First run: if Windows warns about an unknown publisher, choose **More info → Run
anyway** (in-house unsigned tool). If downloaded as a zip, right-click →
Properties → **Unblock** before extracting.
