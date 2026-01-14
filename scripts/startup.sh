#!/bin/bash

################################################################################
# AgentCare - Intelligent Startup Script
# Starts all services with health checks and automatic recovery
################################################################################

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

# Configuration
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LOG_DIR="$PROJECT_DIR/logs"
STARTUP_LOG="$LOG_DIR/startup.log"
HEALTH_CHECK_SCRIPT="$PROJECT_DIR/scripts/health_check.py"
MAX_RETRIES=5
RETRY_DELAY=5

# Ensure log directory exists
mkdir -p "$LOG_DIR"
> "$STARTUP_LOG"  # Clear previous log

################################################################################
# Helper Functions
################################################################################

log_info() {
    echo -e "${BLUE}ℹ${NC} $1" | tee -a "$STARTUP_LOG"
}

log_success() {
    echo -e "${GREEN}✓${NC} $1" | tee -a "$STARTUP_LOG"
}

log_warn() {
    echo -e "${YELLOW}⚠${NC} $1" | tee -a "$STARTUP_LOG"
}

log_error() {
    echo -e "${RED}✗${NC} $1" | tee -a "$STARTUP_LOG"
}

wait_for_service() {
    local service_name=$1
    local check_cmd=$2
    local max_attempts=$3
    local attempt=0

    log_info "Waiting for $service_name to be ready..."

    while [ $attempt -lt $max_attempts ]; do
        if eval "$check_cmd" 2>/dev/null; then
            log_success "$service_name is ready"
            return 0
        fi
        attempt=$((attempt + 1))
        if [ $attempt -lt $max_attempts ]; then
            sleep $RETRY_DELAY
        fi
    done

    log_error "$service_name failed to start after $max_attempts attempts"
    return 1
}

################################################################################
# Phase 1: Pre-startup Checks
################################################################################

phase_precheck() {
    log_info "Phase 1: Pre-startup checks..."

    cd "$PROJECT_DIR"

    # Check Docker
    if ! command -v docker &> /dev/null; then
        log_error "Docker is not installed"
        exit 1
    fi
    log_success "Docker is installed"

    # Check docker-compose
    if ! command -v docker-compose &> /dev/null; then
        log_error "Docker Compose is not installed"
        exit 1
    fi
    log_success "Docker Compose is installed"

    # Check if Docker daemon is running
    if ! docker info &>/dev/null; then
        log_error "Docker daemon is not running"
        log_info "Please start Docker Desktop and try again"
        exit 1
    fi
    log_success "Docker daemon is running"

    # Check for compose file
    if [ ! -f "docker-compose.yml" ]; then
        log_error "docker-compose.yml not found"
        exit 1
    fi
    log_success "docker-compose.yml found"

    # Check .env file
    if [ ! -f ".env" ]; then
        log_warn ".env file not found, using defaults"
    else
        log_success ".env file loaded"
    fi
}

################################################################################
# Phase 2: Cleanup & Preparation
################################################################################

phase_cleanup() {
    log_info "Phase 2: Cleanup & preparation..."

    cd "$PROJECT_DIR"

    # Remove orphaned containers
    log_info "Removing orphaned Docker containers..."
    docker-compose down --remove-orphans || true

    # Prune unused volumes (optional, commented out)
    # docker volume prune -f || true

    log_success "Cleanup complete"
}

################################################################################
# Phase 3: Start Services
################################################################################

phase_start_services() {
    log_info "Phase 3: Starting services..."

    cd "$PROJECT_DIR"

    # Start all services
    log_info "Building and starting Docker containers..."
    docker-compose up -d

    log_success "Docker containers started"

    # Show running containers
    echo ""
    log_info "Running containers:"
    docker-compose ps --services | while read service; do
        status=$(docker-compose ps $service --format "table {{.State}}" 2>/dev/null | tail -1)
        echo "  - $service: $status"
    done
}

################################################################################
# Phase 4: Wait for Services
################################################################################

phase_wait_services() {
    log_info "Phase 4: Waiting for services to be healthy..."

    cd "$PROJECT_DIR"

    # Redis health check
    wait_for_service \
        "Redis" \
        "docker exec agentcare-redis redis-cli ping | grep -q PONG" \
        12 || {
        log_error "Redis failed to start"
        docker-compose logs agentcare-redis
        exit 1
    }

    # FastAPI health check
    wait_for_service \
        "FastAPI Server" \
        "curl -s http://localhost:8000/health &>/dev/null" \
        12 || {
        log_error "FastAPI Server failed to start"
        docker-compose logs agentcare
        exit 1
    }
}

################################################################################
# Phase 5: Health Verification
################################################################################

phase_health_check() {
    log_info "Phase 5: Running comprehensive health checks..."

    if [ ! -f "$HEALTH_CHECK_SCRIPT" ]; then
        log_warn "Health check script not found, skipping"
        return 0
    fi

    # Activate virtual environment if it exists
    if [ -f ".venv/bin/activate" ]; then
        source .venv/bin/activate
    fi

    # Run health check
    if python "$HEALTH_CHECK_SCRIPT"; then
        log_success "All health checks passed"
        return 0
    else
        log_warn "Some health checks failed or skipped (non-critical)"
    fi
}

################################################################################
# Phase 6: Service Details
################################################################################

phase_show_details() {
    log_info "Phase 6: Service information..."

    cd "$PROJECT_DIR"

    echo ""
    echo "────────────────────────────────────────"
    echo "FastAPI Server"
    echo "────────────────────────────────────────"

    # Get API health
    if curl -s http://localhost:8000/health 2>/dev/null > /dev/null; then
        api_info=$(curl -s http://localhost:8000/health)
        log_success "API is responding"
        echo "$api_info" | python -m json.tool 2>/dev/null || echo "$api_info"
    else
        log_warn "API not responding yet"
    fi

    echo ""
    echo "────────────────────────────────────────"
    echo "Redis Cache"
    echo "────────────────────────────────────────"

    if docker exec agentcare-redis redis-cli info stats &>/dev/null; then
        log_success "Redis is running"
        docker exec agentcare-redis redis-cli info stats | grep -E "connected_clients|keyspace" || true
    else
        log_warn "Redis not responding"
    fi

    echo ""
    echo "────────────────────────────────────────"
    echo "Docker Containers"
    echo "────────────────────────────────────────"
    docker-compose ps
}

################################################################################
# Phase 7: Summary
################################################################################

phase_summary() {
    log_info "Phase 7: Startup complete!"

    echo ""
    echo "╔════════════════════════════════════════╗"
    echo "║  ✓ Startup Successful                 ║"
    echo "╚════════════════════════════════════════╝"
    echo ""
    echo "Services are now running!"
    echo ""
    echo "Access points:"
    echo "  • API Docs:    ${BLUE}http://localhost:8000/docs${NC}"
    echo "  • Health:      ${BLUE}http://localhost:8000/health${NC}"
    echo "  • API Base:    ${BLUE}http://localhost:8000${NC}"
    echo ""
    echo "Useful commands:"
    echo "  • View logs:          docker-compose logs -f"
    echo "  • View app logs:      docker-compose logs -f agentcare"
    echo "  • View Redis logs:    docker-compose logs -f agentcare-redis"
    echo "  • Stop services:      docker-compose down"
    echo "  • Run health check:   python scripts/health_check.py"
    echo ""
    echo "Startup log: $STARTUP_LOG"
    echo ""
}

################################################################################
# Error Handling
################################################################################

cleanup_on_error() {
    log_error "Startup failed! Check logs above."
    log_info "To debug:"
    log_info "  • View full logs: docker-compose logs"
    log_info "  • View app logs: docker-compose logs agentcare"
    log_info "  • Check .env file: cat .env"
    echo ""
    echo "Startup log saved to: $STARTUP_LOG"
    exit 1
}

trap cleanup_on_error ERR

################################################################################
# Main Execution
################################################################################

main() {
    echo ""
    echo "╔════════════════════════════════════════╗"
    echo "║  AgentCare - Startup Script          ║"
    echo "╚════════════════════════════════════════╝"
    echo ""

    phase_precheck
    phase_cleanup
    phase_start_services
    phase_wait_services
    phase_health_check
    phase_show_details
    phase_summary
}

main "$@"
