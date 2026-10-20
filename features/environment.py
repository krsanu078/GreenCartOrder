import base64
import io
import logging
import os
import traceback
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def before_all(context):
    context.chrome_options = Options()
    headless = os.getenv("HEADLESS", "1").strip().lower()
    if headless in {"1", "true", "yes"}:
        context.chrome_options.add_argument("--headless=new")
        context.chrome_options.add_argument("--disable-gpu")
        context.chrome_options.add_argument("--no-sandbox")
        context.chrome_options.add_argument("--disable-dev-shm-usage")
    context.chrome_options.add_argument("--window-size=1920,1080")


def before_scenario(context, scenario):
    context.report_capture = {}
    service = ChromeService(ChromeDriverManager().install())
    context.driver = webdriver.Chrome(service=service, options=context.chrome_options)

def before_step(context, step):
    log_stream = io.StringIO()
    log_handler = logging.StreamHandler(log_stream)
    log_handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    )
    root_logger = logging.getLogger()
    previous_level = root_logger.level
    if previous_level > logging.INFO:
        root_logger.setLevel(logging.INFO)
    root_logger.addHandler(log_handler)
    root_logger.info("START %s %s", step.keyword, step.name)
    context.report_capture.update(
        stream=log_stream,
        handler=log_handler,
        logger=root_logger,
        previous_level=previous_level,
    )

def after_step(context, step):
    capture = context.report_capture
    root_logger = capture["logger"]
    log_handler = capture["handler"]
    status = getattr(step.status, "name", str(step.status)).lower()
    duration = getattr(step, "duration", 0.0)
    root_logger.info(
        "END %s %s: %s (%.3fs)", step.keyword, step.name, status, duration
    )
    log_handler.flush()
    root_logger.removeHandler(log_handler)
    root_logger.setLevel(capture["previous_level"])

    log_lines = [
        f"Step: {step.keyword} {step.name}",
        f"Status: {status}",
        f"Duration: {duration:.3f}s",
    ]
    captured_logs = capture["stream"].getvalue().strip()
    if captured_logs:
        log_lines.extend(("", "Python logs:", captured_logs))
    error_message = getattr(step, "error_message", None)
    if error_message:
        log_lines.extend(("", "Failure details:", str(error_message)))
    exception = getattr(step, "exception", None)
    exception_traceback = getattr(step, "exc_traceback", None)
    if exception and exception_traceback:
        log_lines.extend(
            (
                "",
                "Traceback:",
                "".join(
                    traceback.format_exception(
                        type(exception), exception, exception_traceback
                    )
                ),
            )
        )

    try:
        screenshot = context.driver.get_screenshot_as_png()
        encoded_screenshot = base64.b64encode(screenshot).decode("ascii")
        context.attach("image/png", encoded_screenshot)
    except (AttributeError, WebDriverException) as error:
        log_lines.extend(("", f"Screenshot unavailable: {error}"))

    context.attach("text/plain", "\n".join(log_lines))
    log_handler.close()
    capture.clear()

def after_scenario(context, scenario):
    try:
        driver = getattr(context, "driver", None)
        if driver:
            driver.quit()
    except Exception:
        pass
