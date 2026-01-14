#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AgentCare Health Checker
Comprehensive diagnostic tool for system components and connections.
"""

import os
import sys
import time
import json
import socket
import logging
import asyncio
from pathlib import Path
from dotenv import load_dotenv

# Try to import required libraries
try:
    import redis
    import requests
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
except ImportError as e:
    print(f"❌ Ошибка: Отсутствуют зависимости. Запустите 'pip install -r requirements.txt'")
    print(f"Детали: {e}")
    sys.exit(1)

# Configure logging
logging.basicConfig(level=logging.ERROR)

class HealthChecker:
    def __init__(self):
        # Load environment variables
        load_dotenv()
        self.results = {}
        self.config = {
            "google_creds": os.getenv("GOOGLE_CREDENTIALS_FILE", "config/credentials.json"),
            "redis_url": os.getenv("REDIS_URL", "redis://localhost:6379/0"),
            "gemini_api_key": os.getenv("GEMINI_API_KEY"),
            "api_url": f"http://localhost:{os.getenv('SERVER_PORT', '8000')}",
            "supabase_url": os.getenv("DB_HOST"),
        }

    def print_section(self, title):
        print(f"\n{'='*20} {title} {'='*20}")

    def check_env_file(self):
        print("🔍 Проверка .env файла...")
        if Path(".env").exists():
            print("✅ .env файл найден")
            return True
        else:
            print("❌ .env файл ОТСУТСТВУЕТ (используйте .env.template)")
            return False

    def check_google_credentials(self):
        print("🔍 Проверка Google Credentials...")
        creds_path = Path(self.config["google_creds"])
        if not creds_path.exists():
            print(f"❌ Файл {creds_path} не найден")
            return False
        
        try:
            with open(creds_path) as f:
                data = json.load(f)
            print(f"✅ Файл найден: {data.get('client_email')}")
            
            # Test loading
            Credentials.from_service_account_file(str(creds_path))
            print("✅ Формат credentials валидный")
            return True
        except Exception as e:
            print(f"❌ Ошибка в файле {creds_path}: {e}")
            return False

    def check_redis(self):
        print("🔍 Проверка Redis...")
        try:
            r = redis.from_url(self.config["redis_url"])
            if r.ping():
                print(f"✅ Подключение к Redis установлено: {self.config['redis_url']}")
                return True
        except Exception as e:
            print(f"❌ Ошибка Redis: {e}")
            return False

    def check_gemini(self):
        print("🔍 Проверка Gemini API...")
        if not self.config["gemini_api_key"]:
            print("❌ GEMINI_API_KEY не задан")
            return False
        
        # Simple test request (optional, usually key length check is enough for basic health)
        masked_key = self.config["gemini_api_key"][:4] + "..." + self.config["gemini_api_key"][-4:]
        print(f"✅ Ключ найден: {masked_key}")
        return True

    def check_api_server(self):
        print("🔍 Проверка FastAPI Server...")
        try:
            response = requests.get(f"{self.config['api_url']}/health", timeout=5)
            if response.status_code == 200:
                print(f"✅ API сервер работает: {self.config['api_url']}")
                return True
            else:
                print(f"❌ API сервер вернул код {response.status_code}")
                return False
        except Exception:
            print(f"❌ API сервер недоступен по адресу {self.config['api_url']}")
            return False

    def check_docker(self):
        print("🔍 Проверка Docker...")
        # Check if docker socket exists or command works
        try:
            import subprocess
            result = subprocess.run(["docker", "ps"], capture_output=True, text=True)
            if result.returncode == 0:
                print("✅ Docker доступен и запущен")
                return True
            else:
                print("❌ Docker установлен, но не запущен")
                return False
        except FileNotFoundError:
            print("❌ Docker CLI не найден")
            return False

    def run_all(self):
        self.print_section("AgentCare Health Check Report")
        
        checks = [
            ("Environment & Config", self.check_env_file),
            ("Google APIs", self.check_google_credentials),
            ("Redis Content", self.check_redis),
            ("Gemini AI", self.check_gemini),
            ("Docker Services", self.check_docker),
            ("API Endpoint", self.check_api_server),
        ]
        
        passed = 0
        total = len(checks)
        
        for name, check_func in checks:
            status = check_func()
            self.results[name] = "✓" if status else "✗"
            if status: passed += 1
            
        self.print_section("Summary")
        for name, status in self.results.items():
            print(f"{status} {name}")
            
        print(f"\nResult: {passed}/{total} checks passed")
        
        if passed == total:
            print("🎉 Система полностью готова к работе!")
        else:
            print("⚠️ Некоторые компоненты требуют настройки.")

if __name__ == "__main__":
    checker = HealthChecker()
    checker.run_all()
