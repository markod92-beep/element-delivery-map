import sys, os, time
src = sys.argv[1] if len(sys.argv) > 1 else "_src"
maxh = float(os.environ.get("MAX_AGE_HOURS", "48"))
required = ["CMMC_UniqueEvents.xlsx", "CMMC_Dispatch_Routes.xlsx"]
bad = []
for fn in required:
    fp = os.path.join(src, fn)
    if not os.path.exists(fp):
        print(f"MISSING {fn}"); sys.exit(1)
    age = (time.time() - os.path.getmtime(fp)) / 3600.0
    print(f"{fn}: {age:.1f}h old")
    if age > maxh: bad.append(fn)
if bad:
    print(f"STALE (>{maxh}h): {bad} -- source export did not refresh; aborting."); sys.exit(1)
print("freshness ok")
