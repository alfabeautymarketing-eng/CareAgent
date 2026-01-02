import json
import time
from typing import Any, Optional
import redis
from src.utils.config import settings
from src.utils.logger import logger

class CacheService:
    def __init__(self):
        self.redis_client = None
        self._local_cache = {}
        self._local_expiry = {}
        self._connect()

    def _connect(self):
        """Connect to Redis server."""
        try:
            self.redis_client = redis.from_url(settings.redis_url, decode_responses=True)
            self.redis_client.ping()
            logger.info("connected_to_redis")
        except Exception as e:
            logger.warning("redis_connection_failed_using_fallback", error=str(e))
            self.redis_client = None

    def get(self, key: str) -> Optional[Any]:
        """Get value from cache."""
        # Try Redis
        if self.redis_client:
            try:
                data = self.redis_client.get(key)
                if data:
                    return json.loads(data)
            except Exception as e:
                logger.error("redis_get_failed", key=key, error=str(e))

        # Fallback to local
        if key in self._local_cache:
            if time.time() < self._local_expiry.get(key, 0):
                return self._local_cache[key]
            else:
                del self._local_cache[key]
                del self._local_expiry[key]
        
        return None

    def set(self, key: str, value: Any, ttl: int = None):
        """Set value in cache."""
        ttl = ttl or settings.cache_ttl
        
        # Try Redis
        if self.redis_client:
            try:
                self.redis_client.set(key, json.dumps(value), ex=ttl)
            except Exception as e:
                logger.error("redis_set_failed", key=key, error=str(e))

        # Always update local for fast access and fallback
        self._local_cache[key] = value
        self._local_expiry[key] = time.time() + ttl

    def delete(self, key: str):
        """Delete from cache."""
        if self.redis_client:
            try:
                self.redis_client.delete(key)
            except Exception as e:
                logger.error("redis_delete_failed", key=key, error=str(e))
        
        self._local_cache.pop(key, None)
        self._local_expiry.pop(key, None)

cache_service = CacheService()
