import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()
