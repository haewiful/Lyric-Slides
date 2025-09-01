# Install Intel Homebrew
#arch -x86_64 /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Add to ~/.zshrc
#alias ibrew="arch -x86_64 /usr/local/bin/brew"

# Reload zsh
#source ~/.zshrc

# Install Intel Python using Intel Homebrew
#ibrew install python@3.11

## Verify Intel Python
# Check the binary (is the Intel Python has successfully downloaded)
#ls -l /usr/local/bin/python3.11

# Check architecture (Confirm if it's Intel Python)
#file /usr/local/bin/python3.11

# Run Python in Intel mode
#arch -x86_64 /usr/local/bin/python3.11 --version

# Install Intel PyInstaller
#arch -x86_64 /usr/local/bin/pip3.11 install pyinstaller

# Build Intel-compatible executable
arch -x86_64 /usr/local/bin/pyinstaller --onefile --name macIntel_lyrics lyrics.py

# Verify if the binary is Intel-compatible
file dist/macIntel_lyrics
# Expected output: Mach-O 64-bit executable x86_64

