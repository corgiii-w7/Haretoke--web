#!/bin/bash
cd "$(dirname "$0")"
python3 optimize_images.py
echo ""
read -p "Press [Enter] key to exit..."
