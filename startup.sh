#!/bin/bash

# Azure App Service startup script for docx-to-pdf converter
# This script installs system dependencies required for PDF conversion

echo "Installing system dependencies..."

# Update package lists
apt-get update

# Install pandoc and LaTeX for PDF conversion
apt-get install -y pandoc texlive-xetex texlive-fonts-recommended texlive-plain-generic

echo "System dependencies installed successfully"
