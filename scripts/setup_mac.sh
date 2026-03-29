#!/bin/bash
echo "============================================"
echo " App Tools Setup — Mac"
echo "============================================"
echo ""

# Check Git
if ! command -v git &>/dev/null; then
    echo "ERROR: Git is not installed."
    echo "Run this first: xcode-select --install"
    exit 1
fi

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Download: https://www.python.org/downloads/"
    exit 1
fi

# Create tools folder and clone/update repo
echo "Setting up ~/AppTools..."
mkdir -p ~/AppTools

if [ -d ~/AppTools/app-tools ]; then
    echo "Repo already exists — pulling latest..."
    git -C ~/AppTools/app-tools pull
else
    git clone https://github.com/qpx7mxr9/app-tools ~/AppTools/app-tools
fi

# Install Python dependencies
echo "Installing dependencies..."
pip3 install xlwings pandas
pip3 install -e ~/AppTools/app-tools

# Install xlwings Excel add-in
echo "Installing xlwings Excel add-in..."
xlwings addin install

# Create manual update script
cat > ~/AppTools/update.sh << 'EOF'
#!/bin/bash
echo "Updating App Tools..."
git -C ~/AppTools/app-tools pull
echo "Done!"
EOF
chmod +x ~/AppTools/update.sh

# Schedule silent daily auto-pull at 8am via cron
echo "Scheduling daily auto-update at 8:00 AM..."
(crontab -l 2>/dev/null; echo "0 8 * * * git -C ~/AppTools/app-tools pull >> ~/AppTools/update.log 2>&1") | crontab -

echo ""
echo "============================================"
echo " Setup complete!"
echo ""
echo " - Open your Excel workbook and click Run."
echo " - Tools update automatically every morning."
echo " - To update manually: ~/AppTools/update.sh"
echo "============================================"
