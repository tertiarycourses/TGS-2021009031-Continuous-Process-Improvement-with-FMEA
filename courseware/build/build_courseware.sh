#!/usr/bin/env bash
set -euo pipefail

repo_dir="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$repo_dir"

python3 courseware/build/build_activities.py
python3 courseware/build/build_slides.py
python3 courseware/build/build_learner_guide.py
python3 courseware/build/build_lesson_plan.py
if [[ -f courseware/build/build_assessment.py ]]; then
  python3 courseware/build/build_assessment.py
fi

soffice --headless --convert-to pdf --outdir courseware \
  "courseware/Continuous Process Improvement with FMEA-v11.0.pptx" \
  "courseware/LG-Continuous Process Improvement with FMEA-v11.0.docx" \
  "courseware/LP-Continuous Process Improvement with FMEA-v11.0.docx"

python3 courseware/build/inject_toc.py \
  "courseware/LG-Continuous Process Improvement with FMEA-v11.0.docx" \
  "courseware/LG-Continuous Process Improvement with FMEA-v11.0.pdf" 2
python3 courseware/build/inject_toc.py \
  "courseware/LP-Continuous Process Improvement with FMEA-v11.0.docx" \
  "courseware/LP-Continuous Process Improvement with FMEA-v11.0.pdf" 2

soffice --headless --convert-to pdf --outdir courseware \
  "courseware/LG-Continuous Process Improvement with FMEA-v11.0.docx" \
  "courseware/LP-Continuous Process Improvement with FMEA-v11.0.docx"

echo "Courseware build complete."
