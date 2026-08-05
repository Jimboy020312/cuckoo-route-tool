# Setting Up Appium on Windows — Beginner's Guide

This walks through everything needed to get your computer talking to your phone and controlling the Cuckoo+ app automatically. Do these in order — later steps depend on earlier ones working.

**Total time:** probably 45–90 minutes the first time, mostly waiting for installers.

---

## Before you start: what these programs actually do

You don't need to understand these deeply, but it helps to know why you're installing each one:

- **Node.js** — a program that lets you run JavaScript-based tools on your computer. Appium is built on this.
- **Java (JDK)** — Android's own tools (which Appium uses) are built in Java, so they need Java installed to run.
- **Android SDK Platform Tools** — gives you `adb`, the tool that lets your computer send commands to your phone over USB.
- **Appium** — the automation tool that will actually tap buttons and read screens in the Cuckoo+ app for you.
- **Appium Inspector** — a separate app with a visual interface, used only for *looking at* the app's screens to figure out what to tell Appium to click. You won't need it once the script is built.

---

## Step 1: Open Command Prompt

You'll use this window (or one like it) for most of this guide.

1. Press the **Windows key** on your keyboard.
2. Type `cmd`.
3. Click **Command Prompt** when it appears in the search results.

A black window with white text opens. This is where you'll type commands. Keep it open — you'll come back to it repeatedly.

> **Tip:** To paste into Command Prompt, right-click inside the window (there's no Ctrl+V). To copy commands from this guide, select the text normally (Ctrl+C) then right-click to paste into the window.

---

## Step 2: Install Node.js

1. Open your browser and go to **nodejs.org**
2. You'll see two download buttons — click the one labeled **LTS** (this means "long-term support," the stable version).
3. Once downloaded, double-click the installer file (probably in your Downloads folder, named something like `node-v20.x.x-x64.msi`).
4. Click **Next** through the installer screens — the default options are fine. Click **Install**, then **Finish**.
5. **Close and reopen** Command Prompt (important — it needs to reload to recognize the new program).
6. Type this and press Enter:
   ```
   node -v
   ```
7. You should see a version number like `v20.11.0`. If you see an error instead ("not recognized as a command"), restart your computer and try again.

---

## Step 3: Install Java (JDK)

1. Go to **adoptium.net** (this is a free, safe distributor of Java — search "Eclipse Adoptium" if the direct link changes).
2. Download the **JDK 17** version for Windows (x64), the `.msi` installer.
3. Double-click it, click through the installer with default settings, and finish.
4. **Important extra step on Windows** — you need to tell your computer where Java is installed:
   - Press the **Windows key**, type `environment variables`, click **"Edit the system environment variables"**.
   - A window titled "System Properties" opens. Click the **Environment Variables...** button near the bottom.
   - Under the top box ("User variables"), click **New...**
   - Variable name: `JAVA_HOME`
   - Variable value: the folder where Java installed — usually `C:\Program Files\Eclipse Adoptium\jdk-17.x.x-hotspot` (check your actual folder name inside `C:\Program Files\Eclipse Adoptium\`)
   - Click **OK** on all the open windows.
5. Close and reopen Command Prompt, then check:
   ```
   java -version
   ```
   You should see version info, not an error.

---

## Step 4: Install Android SDK Platform Tools (this gives you `adb`)

1. Go to **developer.android.com/tools/releases/platform-tools**
2. Click **"Download SDK Platform-Tools for Windows"**, accept the terms, and it downloads a `.zip` file.
3. Find the downloaded zip in your Downloads folder, right-click it, and choose **Extract All...**
4. Extract it somewhere easy to find and remember — for example, create a folder `C:\platform-tools` and extract there, so you end up with files directly inside `C:\platform-tools` (things like `adb.exe`).
5. Now add this folder to your PATH (this tells Windows "look here when I type a command"):
   - Windows key → type `environment variables` → **"Edit the system environment variables"** again.
   - Click **Environment Variables...**
   - Under "User variables", click on the line that says **Path**, then click **Edit...**
   - Click **New**, and type the folder path exactly: `C:\platform-tools`
   - Click **OK** on everything to close all the windows.
6. Close and reopen Command Prompt, then check:
   ```
   adb version
   ```
   You should see version info like `Android Debug Bridge version 1.0.41`.

---

## Step 5: Install Appium

Back in Command Prompt:

```
npm install -g appium
```

This will print a bunch of text while it downloads — that's normal, wait for it to finish and return to a blank prompt line.

Check it worked:
```
appium -v
```
You should see a version number.

Now install the Android driver Appium needs:
```
appium driver install uiautomator2
```
This also takes a minute or two. Wait for it to finish.

---

## Step 6: Prepare your phone

1. Open **Settings** on your phone.
2. Go to **About phone**.
3. Find **Build number** (sometimes under "Software information").
4. Tap **Build number 7 times quickly**. You'll see a message like "You are now a developer!"
5. Go back to the main Settings screen — you'll now see a new option called **Developer options** (sometimes under "System" or "Additional settings").
6. Open **Developer options** and turn on **USB debugging**.

---

## Step 7: Connect your phone and confirm it's detected

1. Plug your phone into your computer with a USB cable (the same one you'd use to charge it, as long as it can also transfer data — some cheap cables are charge-only).
2. On your phone, a popup should appear: **"Allow USB debugging?"** Tap **Allow** (and tick "always allow from this computer" if you don't want to confirm every time).
3. Your phone might also show a popup asking about the USB connection **mode** — if so, choose **File Transfer** (not "charging only").
4. In Command Prompt, type:
   ```
   adb devices
   ```
5. You should see something like:
   ```
   List of devices attached
   ABC123XYZ    device
   ```
   - If the list is **empty**, unplug and replug the cable, and check the phone screen for the trust popup (it might be waiting for you to tap Allow).
   - If it says `unauthorized` instead of `device`, tap Allow on the phone popup and run the command again.
   - If Windows doesn't recognize the phone at all, you may need to install your phone brand's USB driver (search "[your phone brand] USB driver Windows").

---

## Step 8: Find the Cuckoo+ app's package name

With your phone still connected, type:
```
adb shell pm list packages | findstr cuckoo
```
(Note: Windows uses `findstr` instead of `grep`.)

You should see something like:
```
package:cuckoo.doctress
```
Write this down — you'll need it later. If nothing shows up, try:
```
adb shell pm list packages
```
and scroll through the full list manually looking for anything Cuckoo-related.

---

## Step 9: Start the Appium server

In Command Prompt, type:
```
appium
```
You'll see several lines of text end with something like:
```
[Appium] Appium REST http interface listener started on 0.0.0.0:4723
```

**Leave this window open and running** — this is your Appium server, and it needs to stay on in the background the whole time you're using Inspector or running the script. Don't close this window; open a **new** Command Prompt window for anything else.

---

## Step 10: Install Appium Inspector

This is a separate visual tool — different from the `appium` command you just ran.

1. Go to **github.com/appium/appium-inspector/releases**
2. Scroll to the most recent release, and under "Assets," find the Windows installer — it'll be a file ending in `.exe` with "win" in the name (e.g. `Appium-Inspector-....win.exe`).
3. Download it, then double-click to install, following the prompts.
4. Open **Appium Inspector** from your Start menu once installed.

---

## Step 11: Connect Inspector to your phone

In the Appium Inspector window:

1. Look for a box or tab for **"Capability Builder"** or a raw **JSON input**. Enter these (using the "+" or JSON editor, one at a time if it's a form):

   | Capability | Value |
   |---|---|
   | `platformName` | `Android` |
   | `appium:automationName` | `UiAutomator2` |
   | `appium:appPackage` | `cuckoo.doctress` (or whatever Step 8 gave you) |
   | `appium:appActivity` | `.MainActivity` |
   | `appium:noReset` | `true` |

2. Near the top, there should be fields for **Remote Host** and **Remote Port** — set:
   - Host: `127.0.0.1`
   - Port: `4723`

3. Click **Start Session** (usually a button at the bottom).

**If it fails to connect:** the most common cause is `appium:appActivity` being wrong. Fix by, on your phone, opening the Cuckoo+ app, then in your **second** Command Prompt window (not the one running the server) typing:
```
adb shell dumpsys window | findstr mCurrentFocus
```
This prints something like `cuckoo.doctress/.ui.SplashActivity` — the part after the `/` is your real `appActivity`. Update it in Inspector and try Start Session again.

---

## Step 12: Look around and write down what you need

Once connected, Inspector shows a live picture of your phone's screen.

1. On your **phone**, navigate to the customer list screen in Cuckoo+.
2. In Inspector, click the **refresh icon** (circular arrow, usually top-right of the screen preview) so it re-reads what's currently showing.
3. Click directly on one customer's row **in the Inspector screenshot** (not on your phone). A side panel appears showing details about that element — look for a field called **resource-id** (it'll look like `cuckoo.doctress:id/something`). Write this down — this is your `LIST_ROW_SELECTOR`.
4. On your phone, tap into that customer's record to open the detail view.
5. Back in Inspector, click refresh again, then click on each field you care about (name, address, phone, product) and note its resource-id.
6. On your phone, tap the second tab. Refresh Inspector, click each field there too, and note those resource-ids.
7. Find and note the resource-id of whatever button takes you back to the list.

Keep a Notepad window open and paste each resource-id in as you find it, labeled clearly (e.g. `row = cuckoo.doctress:id/xyz`), so you don't lose track.

---

## Step 13: Close Inspector before running the script

Only one program can hold an Appium session at a time. When you're done gathering locators:

1. In Inspector, click the **session end / quit session** button (often an "X" or door icon).
2. Close the Inspector app.
3. Leave the *server* Command Prompt window (from Step 9) running.

---

## What's next

Open `cuckoo_scraper.py` in a text editor (Notepad works, though something like VS Code is nicer) and paste your real resource-ids into the `CONFIG` section, replacing the placeholder ones. Once that's done, come back and I'll walk you through actually running the script for the first time.

If anything in this guide doesn't match what you're seeing on screen, tell me exactly what step you're on and what you see instead — screenshots described in words are fine, I don't need anything fancy.
