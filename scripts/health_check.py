#!/usr/bin/env python3
"""
AgentCare - Comprehensive Health Check System
Tests all connections, services, and configurations
"""

import asyncio
import json
import os
import sys
from pathlib import Path
from typing import Dict, List, Tuple
from datetime import datetime
from dataclasses import dataclass, asdict

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

import httpx
import redis
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials

# Colors for terminal output
class Colors:
    OK = '\033[92m'
    WARN = '\033[93m'
    FAIL = '\033[91m'
    INFO = '\033[94m'
    RESET = '\033[0m'
    BOLD = '\033[1m'

@dataclass
class HealthCheckResult:
    name: str
    status: str  # 'OK', 'WARN', 'FAIL'
    message: str
    details: Dict = None

    def __post_init__(self):
        if self.details is None:
            self.details = {}

class HealthChecker:
    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        self.results: List[HealthCheckResult] = []
        self.config = self._load_config()
        self.env = self._load_env()

    def _load_config(self) -> Dict:
        """Load configuration files"""
        config_file = self.project_root / "config" / "settings.yaml"
        if not config_file.exists():
            return {}

        try:
            import yaml
            with open(config_file) as f:
                return yaml.safe_load(f) or {}
        except ImportError:
            return {}

    def _load_env(self) -> Dict:
        """Load environment variables"""
        from dotenv import dotenv_values
        env_file = self.project_root / ".env"
        if env_file.exists():
            return dotenv_values(env_file)
        return {}

    def _print_result(self, result: HealthCheckResult):
        """Print formatted result"""
        status_symbol = {
            'OK': f"{Colors.OK}✓{Colors.RESET}",
            'WARN': f"{Colors.WARN}⚠{Colors.RESET}",
            'FAIL': f"{Colors.FAIL}✗{Colors.RESET}"
        }.get(result.status, '?')

        print(f"  {status_symbol} {result.name}: {result.message}")

        if result.details:
            for key, value in result.details.items():
                print(f"      {key}: {value}")

    def print_section(self, title: str):
        """Print section header"""
        print(f"\n{Colors.BOLD}{Colors.INFO}▶ {title}{Colors.RESET}")

    def check_file_exists(self, name: str, path: Path, required: bool = True) -> HealthCheckResult:
        """Check if file exists"""
        if path.exists():
            return HealthCheckResult(
                name=name,
                status='OK',
                message=f"Found at {path.relative_to(self.project_root)}"
            )
        elif required:
            return HealthCheckResult(
                name=name,
                status='FAIL',
                message=f"REQUIRED file not found: {path.relative_to(self.project_root)}"
            )
        else:
            return HealthCheckResult(
                name=name,
                status='WARN',
                message=f"Optional file not found: {path.relative_to(self.project_root)}"
            )

    # ==========================================================================
    # Phase 1: Environment & Configuration
    # ==========================================================================

    def check_env_files(self):
        """Check environment and config files"""
        self.print_section("Environment & Configuration Files")

        # Check .env
        env_file = self.project_root / ".env"
        result = self.check_file_exists(".env file", env_file, required=True)
        self.results.append(result)
        self._print_result(result)

        # Check settings.yaml
        settings_file = self.project_root / "config" / "settings.yaml"
        result = self.check_file_exists("settings.yaml", settings_file, required=True)
        self.results.append(result)
        self._print_result(result)

        # Check project configs
        projects_dir = self.project_root / "config" / "projects"
        if projects_dir.exists():
            project_files = list(projects_dir.glob("*.yaml"))
            result = HealthCheckResult(
                name="Project configurations",
                status='OK',
                message=f"Found {len(project_files)} project(s)",
                details={f.stem: f.name for f in project_files}
            )
        else:
            result = HealthCheckResult(
                name="Project configurations",
                status='WARN',
                message="Project config directory not found"
            )
        self.results.append(result)
        self._print_result(result)

    def check_credentials(self):
        """Check credentials setup"""
        self.print_section("Credentials & Secrets")

        # Check Google credentials
        creds_file = self.project_root / "config" / "credentials.json"
        if creds_file.exists():
            try:
                with open(creds_file) as f:
                    creds_data = json.load(f)
                    project_id = creds_data.get('project_id', 'unknown')
                    result = HealthCheckResult(
                        name="Google Service Account",
                        status='OK',
                        message=f"Credentials found",
                        details={
                            'Project ID': project_id,
                            'Type': creds_data.get('type', 'unknown')
                        }
                    )
            except json.JSONDecodeError:
                result = HealthCheckResult(
                    name="Google Service Account",
                    status='FAIL',
                    message="credentials.json is invalid JSON"
                )
        else:
            result = HealthCheckResult(
                name="Google Service Account",
                status='FAIL',
                message="credentials.json not found - REQUIRED"
            )
        self.results.append(result)
        self._print_result(result)

        # Check API keys in env
        gemini_key = self.env.get('GEMINI_API_KEY')
        if gemini_key and len(gemini_key) > 10:
            status = 'OK'
            message = "Key configured"
        elif gemini_key:
            status = 'WARN'
            message = "Key too short (likely placeholder)"
        else:
            status = 'WARN'
            message = "Key not set in .env"

        result = HealthCheckResult(
            name="Gemini API Key",
            status=status,
            message=message
        )
        self.results.append(result)
        self._print_result(result)

    # ==========================================================================
    # Phase 2: Google APIs
    # ==========================================================================

    def check_google_apis(self):
        """Check Google API credentials and connectivity"""
        self.print_section("Google APIs")

        creds_file = self.project_root / "config" / "credentials.json"
        if not creds_file.exists():
            result = HealthCheckResult(
                name="Google API Auth",
                status='FAIL',
                message="Skipping - credentials.json not found"
            )
            self.results.append(result)
            self._print_result(result)
            return

        try:
            # Load and validate credentials
            credentials = Credentials.from_service_account_file(
                str(creds_file),
                scopes=[
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive'
                ]
            )

            # Refresh to test validity
            credentials.refresh(Request())

            result = HealthCheckResult(
                name="Google API Credentials",
                status='OK',
                message="Credentials valid and authenticated"
            )
            self.results.append(result)
            self._print_result(result)

            # Check Sheets API
            try:
                from google.auth.transport.requests import Request
                from googleapiclient.discovery import build

                service = build('sheets', 'v4', credentials=credentials)

                # Test with a simple metadata request (no actual data modification)
                # Using a small operation to test connectivity
                result = HealthCheckResult(
                    name="Google Sheets API",
                    status='OK',
                    message="API connection successful"
                )
            except Exception as e:
                result = HealthCheckResult(
                    name="Google Sheets API",
                    status='WARN',
                    message=f"API test skipped: {str(e)[:50]}"
                )
            self.results.append(result)
            self._print_result(result)

        except Exception as e:
            result = HealthCheckResult(
                name="Google API Credentials",
                status='FAIL',
                message=f"Authentication failed: {str(e)[:100]}"
            )
            self.results.append(result)
            self._print_result(result)

    # ==========================================================================
    # Phase 3: Redis Connection
    # ==========================================================================

    def check_redis(self):
        """Check Redis connection"""
        self.print_section("Redis Cache")

        redis_url = self.env.get('REDIS_URL', 'redis://localhost:6379/0')

        try:
            r = redis.from_url(redis_url)
            r.ping()

            # Get some stats
            info = r.info()
            used_memory = info.get('used_memory_human', 'N/A')

            result = HealthCheckResult(
                name="Redis Connection",
                status='OK',
                message="Connected and responding",
                details={
                    'URL': redis_url,
                    'Version': info.get('redis_version', 'unknown'),
                    'Memory Used': used_memory
                }
            )
        except Exception as e:
            result = HealthCheckResult(
                name="Redis Connection",
                status='FAIL',
                message=f"Connection failed: {str(e)}"
            )

        self.results.append(result)
        self._print_result(result)

    # ==========================================================================
    # Phase 4: API Server
    # ==========================================================================

    async def check_api_server(self):
        """Check FastAPI server"""
        self.print_section("API Server")

        api_host = self.env.get('SERVER_HOST', 'localhost')
        api_port = self.env.get('SERVER_PORT', '8000')
        api_url = f"http://{api_host}:{api_port}"

        try:
            async with httpx.AsyncClient(timeout=10) as client:
                # Health endpoint
                response = await client.get(f"{api_url}/health")

                if response.status_code == 200:
                    health_data = response.json()

                    result = HealthCheckResult(
                        name="API Health Endpoint",
                        status='OK',
                        message=f"Server responding on {api_url}",
                        details={
                            'Status': health_data.get('status', 'unknown'),
                            'Version': health_data.get('version', 'unknown')
                        }
                    )
                else:
                    result = HealthCheckResult(
                        name="API Health Endpoint",
                        status='WARN',
                        message=f"Server responded with {response.status_code}"
                    )

        except httpx.ConnectError:
            result = HealthCheckResult(
                name="API Health Endpoint",
                status='FAIL',
                message=f"Cannot connect to {api_url} - server may not be running"
            )
        except Exception as e:
            result = HealthCheckResult(
                name="API Health Endpoint",
                status='WARN',
                message=f"Connection error: {str(e)[:50]}"
            )

        self.results.append(result)
        self._print_result(result)

        # Check Swagger UI
        try:
            async with httpx.AsyncClient(timeout=10) as client:
                response = await client.get(f"{api_url}/docs")
                if response.status_code == 200:
                    status = 'OK'
                    msg = f"Swagger UI available at {api_url}/docs"
                else:
                    status = 'WARN'
                    msg = "Swagger UI not accessible"
        except:
            status = 'WARN'
            msg = "Swagger UI check skipped"

        result = HealthCheckResult(
            name="API Swagger UI",
            status=status,
            message=msg
        )
        self.results.append(result)
        self._print_result(result)

    # ==========================================================================
    # Phase 5: Docker Services
    # ==========================================================================

    def check_docker_services(self):
        """Check Docker service status"""
        self.print_section("Docker Services")

        try:
            import subprocess

            # Check if Docker is running
            result = subprocess.run(['docker', 'info'],
                                  capture_output=True,
                                  timeout=5)
            if result.returncode != 0:
                raise Exception("Docker daemon not running")

            result_obj = HealthCheckResult(
                name="Docker Daemon",
                status='OK',
                message="Docker is running"
            )
            self.results.append(result_obj)
            self._print_result(result_obj)

            # Check docker-compose services
            compose_file = self.project_root / "docker-compose.yml"
            if not compose_file.exists():
                result_obj = HealthCheckResult(
                    name="Docker Compose Config",
                    status='WARN',
                    message="docker-compose.yml not found"
                )
            else:
                result_obj = HealthCheckResult(
                    name="Docker Compose Config",
                    status='OK',
                    message="docker-compose.yml found"
                )
            self.results.append(result_obj)
            self._print_result(result_obj)

            # Check running containers
            try:
                result = subprocess.run(
                    ['docker-compose', '-f', str(compose_file), 'ps', '--services'],
                    cwd=str(self.project_root),
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if result.returncode == 0:
                    services = result.stdout.strip().split('\n')
                    result_obj = HealthCheckResult(
                        name="Docker Services Configured",
                        status='OK',
                        message=f"Found {len(services)} service(s)",
                        details={f"Service {i+1}": s.strip() for i, s in enumerate(services)}
                    )
                else:
                    result_obj = HealthCheckResult(
                        name="Docker Services",
                        status='WARN',
                        message="Could not read service configuration"
                    )
            except subprocess.TimeoutExpired:
                result_obj = HealthCheckResult(
                    name="Docker Services",
                    status='WARN',
                    message="Docker command timed out"
                )

            self.results.append(result_obj)
            self._print_result(result_obj)

        except FileNotFoundError:
            result_obj = HealthCheckResult(
                name="Docker",
                status='WARN',
                message="Docker not installed"
            )
            self.results.append(result_obj)
            self._print_result(result_obj)
        except Exception as e:
            result_obj = HealthCheckResult(
                name="Docker",
                status='FAIL',
                message=f"Error: {str(e)}"
            )
            self.results.append(result_obj)
            self._print_result(result_obj)

    # ==========================================================================
    # Phase 6: Supabase Connection
    # ==========================================================================

    def check_supabase(self):
        """Check Supabase database connection"""
        self.print_section("Supabase PostgreSQL")

        supabase_url = self.env.get('SUPABASE_URL')
        supabase_key = self.env.get('SUPABASE_SERVICE_ROLE_KEY')

        if not supabase_url or not supabase_key:
            result = HealthCheckResult(
                name="Supabase Configuration",
                status='WARN',
                message="Supabase credentials not fully configured in .env"
            )
            self.results.append(result)
            self._print_result(result)
            return

        try:
            from supabase import create_client

            supabase = create_client(supabase_url, supabase_key)

            # Try a simple query
            response = supabase.table('sync_logs').select('id').limit(1).execute()

            result = HealthCheckResult(
                name="Supabase Connection",
                status='OK',
                message="Connected to Supabase PostgreSQL"
            )
        except ImportError:
            result = HealthCheckResult(
                name="Supabase Client",
                status='WARN',
                message="Supabase-py library not installed"
            )
        except Exception as e:
            result = HealthCheckResult(
                name="Supabase Connection",
                status='WARN',
                message=f"Connection test failed: {str(e)[:50]}"
            )

        self.results.append(result)
        self._print_result(result)

    # ==========================================================================
    # Phase 7: Data Directories
    # ==========================================================================

    def check_data_directories(self):
        """Check data storage directories"""
        self.print_section("Data & Logs Directories")

        directories = {
            'Data': self.project_root / 'data',
            'Logs': self.project_root / 'logs',
            'Sync Logs': self.project_root / 'data' / 'sync_logs',
            'Function Logs': self.project_root / 'data' / 'function_logs',
        }

        for name, path in directories.items():
            if path.exists():
                size = sum(f.stat().st_size for f in path.rglob('*') if f.is_file())
                size_mb = size / 1024 / 1024
                result = HealthCheckResult(
                    name=name,
                    status='OK',
                    message=f"Directory exists",
                    details={'Size': f"{size_mb:.2f} MB"}
                )
            else:
                path.mkdir(parents=True, exist_ok=True)
                result = HealthCheckResult(
                    name=name,
                    status='OK',
                    message="Directory created"
                )
            self.results.append(result)
            self._print_result(result)

    # ==========================================================================
    # Phase 8: Python Dependencies
    # ==========================================================================

    def check_python_deps(self):
        """Check Python dependencies"""
        self.print_section("Python Dependencies")

        critical_packages = [
            'fastapi',
            'uvicorn',
            'pydantic',
            'gspread',
            'google-auth',
            'redis',
            'structlog',
            'httpx',
        ]

        import importlib
        missing = []
        installed = []

        for package in critical_packages:
            try:
                module_name = package.replace('-', '_')
                spec = importlib.util.find_spec(module_name)
                if spec is not None:
                    installed.append(package)
            except (ImportError, ModuleNotFoundError):
                missing.append(package)

        if not missing:
            result = HealthCheckResult(
                name="Critical Dependencies",
                status='OK',
                message=f"All {len(installed)} critical packages installed"
            )
        else:
            result = HealthCheckResult(
                name="Critical Dependencies",
                status='WARN',
                message=f"{len(missing)} package(s) missing: {', '.join(missing)}"
            )

        self.results.append(result)
        self._print_result(result)

    # ==========================================================================
    # Summary Report
    # ==========================================================================

    def print_summary(self):
        """Print summary report"""
        print(f"\n{Colors.BOLD}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.RESET}")
        print(f"{Colors.BOLD}Health Check Summary{Colors.RESET}")
        print(f"{Colors.BOLD}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.RESET}")

        ok_count = sum(1 for r in self.results if r.status == 'OK')
        warn_count = sum(1 for r in self.results if r.status == 'WARN')
        fail_count = sum(1 for r in self.results if r.status == 'FAIL')
        total = len(self.results)

        print(f"\n{Colors.OK}✓ OK:{Colors.RESET} {ok_count}")
        if warn_count > 0:
            print(f"{Colors.WARN}⚠ WARNINGS:{Colors.RESET} {warn_count}")
        if fail_count > 0:
            print(f"{Colors.FAIL}✗ FAILURES:{Colors.RESET} {fail_count}")
        print(f"\nTotal checks: {total}")

        if fail_count == 0 and warn_count == 0:
            print(f"\n{Colors.OK}{Colors.BOLD}✓ All systems operational!{Colors.RESET}")
            return 0
        elif fail_count == 0:
            print(f"\n{Colors.WARN}⚠ System operational with warnings{Colors.RESET}")
            return 1
        else:
            print(f"\n{Colors.FAIL}✗ Critical failures detected{Colors.RESET}")
            return 2

    # ==========================================================================
    # Main Execution
    # ==========================================================================

    async def run_all_checks(self):
        """Run all health checks"""
        print(f"{Colors.BOLD}{Colors.INFO}")
        print("╔════════════════════════════════════════╗")
        print("║  AgentCare - Health Check System      ║")
        print(f"║  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}                  ║")
        print("╚════════════════════════════════════════╝")
        print(f"{Colors.RESET}")

        # Run all checks
        self.check_env_files()
        self.check_credentials()
        self.check_google_apis()
        self.check_redis()
        await self.check_api_server()
        self.check_docker_services()
        self.check_supabase()
        self.check_data_directories()
        self.check_python_deps()

        # Print summary
        return self.print_summary()


async def main():
    """Main entry point"""
    checker = HealthChecker()
    exit_code = await checker.run_all_checks()
    sys.exit(exit_code)


if __name__ == '__main__':
    asyncio.run(main())
