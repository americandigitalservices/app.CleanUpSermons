# macOS Setup Guide: Homebrew & Python 3

This guide walks through installing **Homebrew** and **Python 3** on a fresh Mac so the Sermon Outline Cleaner (and its AppleScript launcher) can run automatically[cite: 17, 18].

---

## Step 1: Install Homebrew

[Homebrew](https://brew.sh/) is the standard package manager for macOS.

1. Open **Terminal** (press `Cmd + Space`, type `Terminal`, and press **Enter**).
2. Paste the official Homebrew installation command and press **Enter**:
   ```bash
   /bin/bash -c "$(curl -fsSL [https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh](https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh))"

```

3. Follow the on-screen prompts:
* Enter your Mac login password when prompted (characters will not appear on screen as you type).
* If prompted to install the **Xcode Command Line Tools**, press **Enter** and wait for the installation to finish.



### Add Homebrew to Your PATH (Apple Silicon Macs)

At the end of the installation on M1/M2/M3/M4 Macs, Homebrew may display a **"Next steps"** section asking you to add `brew` to your shell profile. Run these two commands if prompted:

```bash
echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
eval "$(/opt/homebrew/bin/brew shellenv)"

```

Verify Homebrew is installed and ready:

```bash
brew --version

```

---

## Step 2: Install Python 3 via Homebrew

Once Homebrew is installed, install the latest version of Python 3:

```bash
brew install python

```

---

## Step 3: Verify Python 3 & pip

Confirm that Homebrew's `python3` and `pip3` are active:

```bash
python3 --version
which python3

```

* On **Apple Silicon Macs**, `which python3` should output `/opt/homebrew/bin/python3`.


* On **Intel Macs**, `which python3` should output `/usr/local/bin/python3`.



---

## Next Steps

Once `python3` is installed via Homebrew, you are completely done with system setup!

* **Using the AppleScript App (`Launcher`):** Simply double-click the app or drag-and-drop a sermon file onto it. It will automatically detect your Homebrew `python3`, create the local `venv` folder, and install `requirements.txt` on its first run.


* **Using the Command Line:** Follow the manual virtual environment (`venv`) instructions in `README.md`.

