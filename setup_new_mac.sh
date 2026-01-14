#!/bin/bash

################################################################################
# AgentCare - New MacBook Setup & Migration Script
# This script automates the complete setup on a new device
# Usage: bash setup_new_mac.sh [--skip-docker] [--skip-credentials] [--source-dir /path]
################################################################################

set -e  # Exit on error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Configuration
REPO_URL="https://github.com/yourusername/AgentCare.git"  # Update with your repo
PROJECT_DIR="$HOME/Desktop/AgentCare"
BACKUP_DIR="$HOME/.agentcare_backup"
SETUP_LOG="$PROJECT_DIR/setup.log"

# Flags
SKIP_DOCKER=false
SKIP_CREDENTIALS=false
SOURCE_DEVICE_DIR=""
ASSUME_YES=false

################################################################################
# Helper Functions
################################################################################

log_info() {
    echo -e "${BLUE}ℹ${NC} $1" | tee -a "$SETUP_LOG"
}

log_success() {
    echo -e "${GREEN}✓${NC} $1" | tee -a "$SETUP_LOG"
}

log_warn() {
    echo -e "${YELLOW}⚠${NC} $1" | tee -a "$SETUP_LOG"
}

log_error() {
    echo -e "${RED}✗${NC} $1" | tee -a "$SETUP_LOG"
}

confirm() {
    if [ "$ASSUME_YES" = true ]; then
        return 0
    fi
    local prompt="$1"
    local response
    read -p "$(echo -e ${YELLOW}$prompt${NC})" response
    [[ "$response" =~ ^[Yy]$ ]]
}

################################################################################
# Phase 1: Prerequisites Check
################################################################################

phase_prerequisites() {
    log_info "Phase 1: Checking prerequisites..."

    # Check if running on macOS
    if [[ "$OSTYPE" != "darwin"* ]]; then
        log_error "This script is designed for macOS"
        exit 1
    fi

    log_success "Running on macOS"

    # Check/install Homebrew
    if ! command -v brew &> /dev/null; then
        log_warn "Homebrew not found, installing..."
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
        log_success "Homebrew installed"
    else
        log_success "Homebrew is installed"
    fi

    # Check/install Command Line Tools
    if ! xcode-select -p > /dev/null 2>&1; then
        log_warn "Xcode Command Line Tools not found, installing..."
        xcode-select --install
        log_success "Xcode Command Line Tools installed"
    else
        log_success "Xcode Command Line Tools are installed"
    fi

    # Check/install git
    if ! command -v git &> /dev/null; then
        log_warn "Git not found, installing via brew..."
        brew install git
        log_success "Git installed"
    else
        log_success "Git is installed: $(git --version)"
    fi

    # Check/install Python 3.11+
    if ! command -v python3 &> /dev/null; then
        log_warn "Python3 not found, installing via brew..."
        brew install python@3.11
        log_success "Python 3.11 installed"
    else
        local py_version=$(python3 --version | awk '{print $2}')
        log_success "Python is installed: $py_version"
    fi

    # Check/install Poetry
    if ! command -v poetry &> /dev/null; then
        log_warn "Poetry not found, installing..."
        curl -sSL https://install.python-poetry.org | python3 -
        export PATH="$HOME/.local/bin:$PATH"
        log_success "Poetry installed"
    else
        log_success "Poetry is installed: $(poetry --version)"
    fi

    log_success "All prerequisites are met"
}

################################################################################
# Phase 2: Repository Setup
################################################################################

phase_repository() {
    log_info "Phase 2: Repository setup..."

    if [ -d "$PROJECT_DIR" ]; then
        if confirm "Project directory already exists at $PROJECT_DIR. Update from git? (y/n) "; then
            cd "$PROJECT_DIR"
            log_info "Pulling latest code..."
            git fetch origin
            git reset --hard origin/main
            log_success "Repository updated"
        else
            log_warn "Skipping repository update"
        fi
    else
        log_info "Cloning repository..."
        git clone "$REPO_URL" "$PROJECT_DIR"
        cd "$PROJECT_DIR"
        log_success "Repository cloned to $PROJECT_DIR"
    fi

    cd "$PROJECT_DIR"
    log_success "Current branch: $(git rev-parse --abbrev-ref HEAD)"
    log_success "Latest commit: $(git log -1 --pretty=format:'%h - %s')"
}

################################################################################
# Phase 3: Python Dependencies
################################################################################

phase_python_deps() {
    log_info "Phase 3: Installing Python dependencies..."

    cd "$PROJECT_DIR"

    # Create virtual environment if needed
    if [ ! -d ".venv" ]; then
        log_info "Creating virtual environment..."
        python3 -m venv .venv
        log_success "Virtual environment created"
    fi

    # Activate virtual environment
    source .venv/bin/activate

    # Upgrade pip
    log_info "Upgrading pip..."
    python -m pip install --upgrade pip setuptools wheel

    # Install Poetry dependencies
    log_info "Installing Poetry dependencies..."
    poetry install

    log_success "All Python dependencies installed"
}

################################################################################
# Phase 4: Credentials & Configuration
################################################################################

phase_credentials() {
    log_info "Phase 4: Setting up credentials and configuration..."

    if [ "$SKIP_CREDENTIALS" = true ]; then
        log_warn "Skipping credentials setup (--skip-credentials flag)"
        return 0
    fi

    cd "$PROJECT_DIR"

    # Check for existing credentials
    if [ -f "config/credentials.json" ]; then
        log_success "Google credentials already exist"
    else
        log_warn "Google credentials not found"

        if [ -n "$SOURCE_DEVICE_DIR" ] && [ -f "$SOURCE_DEVICE_DIR/config/credentials.json" ]; then
            log_info "Copying credentials from source device..."
            cp "$SOURCE_DEVICE_DIR/config/credentials.json" "config/credentials.json"
            log_success "Credentials copied"
        else
            log_error "Credentials not found. Please provide credentials.json in config/ directory"
            log_info "You can:"
            log_info "  1. Copy from another device: setup_new_mac.sh --source-dir /path/to/old/device"
            log_info "  2. Place manually: config/credentials.json"
            exit 1
        fi
    fi

    # Check .env file
    if [ -f ".env" ]; then
        log_success "Environment file exists"
    else
        log_warn "Creating .env from template..."
        if [ -f ".env.example" ]; then
            cp ".env.example" ".env"
            log_success ".env created from template - EDIT WITH YOUR VALUES"
        else
            log_error ".env.example not found"
            exit 1
        fi
    fi

    # Check project configurations
    if [ -d "config/projects" ]; then
        local count=$(ls config/projects/*.yaml 2>/dev/null | wc -l)
        log_success "Found $count project configuration files"
    fi

    # Check sync rules
    if [ -d "config/rules" ]; then
        local count=$(ls config/rules/*.yaml 2>/dev/null | wc -l)
        log_success "Found $count sync rule files"
    fi

    log_success "Credentials and configuration ready"
}

################################################################################
# Phase 5: Docker Setup
################################################################################

phase_docker() {
    log_info "Phase 5: Setting up Docker..."

    if [ "$SKIP_DOCKER" = true ]; then
        log_warn "Skipping Docker setup (--skip-docker flag)"
        return 0
    fi

    # Check if Docker is installed
    if ! command -v docker &> /dev/null; then
        log_warn "Docker not found"
        if confirm "Install Docker Desktop? (y/n) "; then
            log_info "Installing Docker Desktop via brew..."
            brew install --cask docker
            log_warn "Please start Docker Desktop from Applications folder"
            log_warn "Waiting for Docker to start..."
            sleep 10
        else
            log_error "Docker is required for production setup"
            return 1
        fi
    fi

    # Wait for Docker daemon
    local max_attempts=30
    local attempt=0
    while ! docker info &>/dev/null; do
        if [ $attempt -ge $max_attempts ]; then
            log_error "Docker daemon did not start"
            return 1
        fi
        attempt=$((attempt + 1))
        sleep 1
    done

    log_success "Docker is running"

    # Check docker-compose
    if ! command -v docker-compose &> /dev/null; then
        log_info "Installing Docker Compose..."
        brew install docker-compose
        log_success "Docker Compose installed"
    else
        log_success "Docker Compose is installed: $(docker-compose --version)"
    fi
}

################################################################################
# Phase 6: Services Startup
################################################################################

phase_startup() {
    log_info "Phase 6: Starting services..."

    cd "$PROJECT_DIR"

    # Make startup script executable
    if [ -f "scripts/startup.sh" ]; then
        chmod +x scripts/startup.sh
        log_info "Running startup script..."
        bash scripts/startup.sh
    else
        log_warn "startup.sh not found, starting manually..."

        # Start Redis
        log_info "Starting Redis..."
        docker-compose up -d agentcare-redis

        # Start FastAPI
        log_info "Starting FastAPI server..."
        docker-compose up -d agentcare
    fi

    log_success "Services started"
}

################################################################################
# Phase 7: Health Checks
################################################################################

phase_health_check() {
    log_info "Phase 7: Running comprehensive health checks..."

    cd "$PROJECT_DIR"

    source .venv/bin/activate

    if [ -f "scripts/health_check.py" ]; then
        log_info "Running health_check.py..."
        python scripts/health_check.py
    else
        log_error "health_check.py not found"
        return 1
    fi

    log_success "Health checks complete"
}

################################################################################
# Phase 8: Verification
################################################################################

phase_verification() {
    log_info "Phase 8: Verifying setup..."

    local errors=0

    # Check API endpoint
    log_info "Testing API endpoint..."
    if curl -s http://localhost:8000/health &>/dev/null; then
        log_success "API is responding"
    else
        log_error "API is not responding"
        errors=$((errors + 1))
    fi

    # Check Redis
    log_info "Testing Redis connection..."
    if docker exec agentcare-redis redis-cli ping &>/dev/null | grep -q PONG; then
        log_success "Redis is responding"
    else
        log_error "Redis is not responding"
        errors=$((errors + 1))
    fi

    # Check Docker containers
    log_info "Checking Docker containers..."
    local running=$(docker-compose ps --services --filter "status=running" | wc -l)
    local total=$(docker-compose config --services | wc -l)
    if [ "$running" -eq "$total" ]; then
        log_success "All Docker containers are running ($running/$total)"
    else
        log_warn "Some containers are not running ($running/$total)"
    fi

    if [ $errors -eq 0 ]; then
        log_success "All verifications passed"
        return 0
    else
        log_error "Some verifications failed"
        return 1
    fi
}

################################################################################
# Phase 9: Summary & Next Steps
################################################################################

phase_summary() {
    log_info "Phase 9: Setup complete!"

    echo ""
    echo "========================================"
    echo "AgentCare Setup Summary"
    echo "========================================"
    log_success "✓ Prerequisites checked"
    log_success "✓ Repository cloned/updated"
    log_success "✓ Python dependencies installed"
    log_success "✓ Credentials configured"
    log_success "✓ Docker setup complete"
    log_success "✓ Services started"
    log_success "✓ Health checks passed"
    echo ""
    echo "Next steps:"
    echo "  1. Dashboard: http://localhost:8000/docs"
    echo "  2. Health check: http://localhost:8000/health"
    echo "  3. Edit .env file with your actual values if needed"
    echo "  4. Check logs: tail -f $PROJECT_DIR/logs/app.log"
    echo ""
    echo "To view container logs:"
    echo "  docker-compose logs -f"
    echo ""
    echo "Setup log saved to: $SETUP_LOG"
    echo "========================================"
}

################################################################################
# Error Handling
################################################################################

cleanup_on_error() {
    log_error "Setup failed. Check $SETUP_LOG for details."
    exit 1
}

trap cleanup_on_error ERR

################################################################################
# Parse Arguments
################################################################################

while [[ $# -gt 0 ]]; do
    case $1 in
        --skip-docker)
            SKIP_DOCKER=true
            shift
            ;;
        --skip-credentials)
            SKIP_CREDENTIALS=true
            shift
            ;;
        --source-dir)
            SOURCE_DEVICE_DIR="$2"
            shift 2
            ;;
        --yes)
            ASSUME_YES=true
            shift
            ;;
        *)
            log_error "Unknown option: $1"
            exit 1
            ;;
    esac
done

################################################################################
# Main Execution
################################################################################

main() {
    # Initialize log file
    mkdir -p "$(dirname "$SETUP_LOG")"
    > "$SETUP_LOG"  # Clear previous log

    echo ""
    echo "╔════════════════════════════════════════╗"
    echo "║  AgentCare - New MacBook Setup        ║"
    echo "╚════════════════════════════════════════╝"
    echo ""

    phase_prerequisites
    phase_repository
    phase_python_deps
    phase_credentials
    phase_docker
    phase_startup
    phase_health_check
    phase_verification
    phase_summary
}

main "$@"
