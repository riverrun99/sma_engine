# Changelog — SMA Engine Updates

## 2026-06-01 / 2026-06-02

### Muted Tickers
- Added 150+ tickers to `muted_tickers.txt` across several sessions
- Categories muted: bond ETFs (HYG, BNDW, DBA, SOYB), tiny REITs (PDM, DEI, FBRT, NTST, OLP, GOOD, CHCT, SQFT, BRT, LAND), biotech microcaps (BCAB, FATE, CODX, GNSS, SABS, SPRO, GNTA, NKTX, ARDX, ACRS, TCRT, SVRA, TBPH, ELAB, CLAR), small community banks (RVSB, FNWB, ALEC, PROV, KELYA, SHBI), China/EM small caps (NIO, WB, SEA, ITUB, EWH), noise ETFs (SZK, MSFD, SCC, SMDD, METD, AMDD, CMTG, MSFO, SPDN, TRTX, GGLS, AAPD), illiquid micro caps (IVDA, XOMA, EVMO, SPOK, ZNTL, TTEC)

### Custom Tickers
- Added 88 tickers to `custom_tickers.txt` from personal watchlists
- Includes: leveraged ETFs (TYO, UBT, UPW, UBOT, URAA, WTIU, XOMX, TSMX), watchlist names (AMC, KITT, GFAI, APLD, SYM, VSTL, WTI), reinsurance ADRs (SSREY, HVRRY), composites/aerospace (HXL, TDY), nuclear/energy (CCJ, CEG, LEU, BWXT, PEG), and many others

### Engine Config
- `ENGINE_TOP_N` set to 2000 (was 50)
- `ENGINE_TIMEFRAMES` confirmed: 1m, 5m, 15m, 30m, 1h, 2h, 4h, 1d, 1w, 1mo
- `ENGINE_LOOKBACK` = 130 bars
- `ENGINE_BACKTEST_EVERY` = 288 cycles (auto-backtest every 24 hours)

### local_writer.py — New Features
- **Performance tab** in `signals_current.xlsx`: tracks price change since each ticker's first appearance in ranked output. Persisted via `output/price_tracker.json`.
- **Ranked log** (`output/ranked_log.csv`): append-only, one row per ranked entry per cycle. Full historical record of every ranked signal.
- **Snapshots** (`output/snapshots/snapshot_YYYY-MM-DD_HH-MM-SS.csv`): one timestamped file per scan cycle listing every unique ticker with timeframe, outfit, hits, convergence, score.

### backtest.py
- Fixed `ZeroDivisionError` in `evaluate_signal` when candle open price is zero

### run_backtest.py — New File
- Standalone backtest runner: reads current `signals_current.xlsx`, loads candle cache, runs walk-forward backtest on top 100 signals
- Prints summary table: ticker, timeframe, outfit (SMA periods), trades, win rate, avg return, Sharpe
- Saves timestamped CSV to `output/backtest_YYYY-MM-DD_HH-MM-SS.csv`
- Usage: `docker exec e47_engine python /app/run_backtest.py`
- Options: `--method cpcv`, `--horizon N`, `--top N`

### daemon.py
- Wired `ENGINE_BACKTEST_EVERY` env var to auto-run backtest every N cycles
- Added `backtest_every` logic in main loop

### README.md
- Updated placeholder URLs from YOUR_USERNAME to riverrun99
