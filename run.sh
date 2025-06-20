#!/bin/bash
# Save as ~/run-cool-fixed.sh

# Clean environment
unset PYTHONPATH
unset PYTHONHOME

# Check actual structure
PYTHON_DIR="/home/humphry/Desktop/sdocs/code/instdir/program/python-core-3.10.17"
PYTHON_LIB_DIR=$(find $PYTHON_DIR -name "encodings" -type d | head -n1 | xargs dirname)

if [ -z "$PYTHON_LIB_DIR" ]; then
  echo "ERROR: Could not find encodings module directory!"
  exit 1
fi

echo "Found Python library directory: $PYTHON_LIB_DIR"

# Set proper environment
export PYTHONHOME=$PYTHON_DIR
export PYTHONPATH=$PYTHON_DIR:$PYTHON_LIB_DIR
export PYTHONEXECUTABLE=/home/humphry/Desktop/sdocs/code/instdir/program/python

# Force the system to use the correct Python binary
echo "Creating Python symlink..."
mkdir -p ~/bin
ln -sf /home/humphry/Desktop/sdocs/code/instdir/program/python ~/bin/python3
export PATH=~/bin:$PATH

# Test Python setup
echo "Testing Python setup..."
~/bin/python3 -c "import sys; print('Python version:', sys.version); import encodings; print('Encodings module found!')"

# Run make
cd ~/Desktop/sdocs/code/collabora-online
make run