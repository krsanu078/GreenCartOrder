from pathlib import Path
import subprocess
import sys


PROJECT_ROOT = Path(__file__).resolve().parent
FEATURES_IN_ORDER = (
    "features/search_product.feature",
    "features/add_to_cart.feature",
    "features/cart_checkout.feature",
    "features/promo_code.feature",
    "features/place_order.feature",
)


def main():
    unit_tests = [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"]
    unit_test_result = subprocess.call(unit_tests, cwd=PROJECT_ROOT)
    if unit_test_result:
        return unit_test_result

    command = [sys.executable, "-m", "behave", *FEATURES_IN_ORDER]
    return subprocess.call(command, cwd=PROJECT_ROOT)


if __name__ == "__main__":
    raise SystemExit(main())
