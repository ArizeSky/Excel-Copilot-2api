import os
import unittest


class ConfigTests(unittest.TestCase):
    def test_defaults(self):
        env_keys = ["PROXY_PORT", "PROXY_API_KEY", "CHROME_PROFILE_DIR", "PROXY_DRIVER"]
        saved = {k: os.environ.pop(k, None) for k in env_keys}
        try:
            import importlib
            import config
            importlib.reload(config)
            self.assertEqual(config.PROXY_PORT, 12803)
            self.assertEqual(config.PROXY_API_KEY, "")
            self.assertIn("chrome-profile", config.CHROME_PROFILE_DIR)
            self.assertEqual(config.PROXY_DRIVER, "auto")
        finally:
            for k, v in saved.items():
                if v is not None:
                    os.environ[k] = v

    def test_env_override(self):
        os.environ["PROXY_PORT"] = "9999"
        os.environ["PROXY_API_KEY"] = "sk-test-123"
        os.environ["CHROME_PROFILE_DIR"] = "/tmp/my-profile"
        os.environ["PROXY_DRIVER"] = "linux"
        try:
            import importlib
            import config
            importlib.reload(config)
            self.assertEqual(config.PROXY_PORT, 9999)
            self.assertEqual(config.PROXY_API_KEY, "sk-test-123")
            self.assertEqual(config.CHROME_PROFILE_DIR, "/tmp/my-profile")
            self.assertEqual(config.PROXY_DRIVER, "linux")
        finally:
            for k in ["PROXY_PORT", "PROXY_API_KEY", "CHROME_PROFILE_DIR", "PROXY_DRIVER"]:
                os.environ.pop(k, None)


if __name__ == "__main__":
    unittest.main()
