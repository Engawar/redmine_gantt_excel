rem redmine_gantt_excel

echo 全ステータス、日単位ガント
python redmine_gantt_excel.py ^
  --status-id "*" ^
  --timeline-mode day ^
  --output redmine_gantt.xlsx

echo 週単位で圧縮、期間を固定
python redmine_gantt_excel.py ^
  --timeline-mode week ^
  --from-date 2026-04-01 ^
  --to-date 2026-06-30 ^
  --output redmine_gantt_weekly.xlsx
