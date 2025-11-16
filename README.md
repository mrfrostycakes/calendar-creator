# Calendar Creator

A Python script to generate beautiful year-at-a-glance calendars in PowerPoint format with customizable color themes.

## Features

- 📅 Generates a full year calendar with all 12 months on a single slide
- 🎨 12 color themes to choose from (cool tones, warm tones, and neutrals)
- 📐 Widescreen 16:9 format (13.333" x 7.5")
- ✨ Clean, professional layout with proper spacing
- 🎯 Dynamic row calculation - only shows the rows each month needs
- 🔤 Arial 12pt font throughout

## Installation

1. Clone this repository:
```bash
git clone https://github.com/mrfrostycakes/calendar-creator.git
cd calendar-creator
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
.\venv\Scripts\Activate.ps1  # Windows PowerShell
# or
source venv/bin/activate     # macOS/Linux
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the script with a year:
```bash
python make_calendar.py 2026
```

You'll be prompted to choose a color theme:

**Cool Tones:**
- 1. Blue (Sky Blue)
- 2. Navy (Professional Blue)
- 3. Teal
- 4. Purple

**Warm Tones:**
- 5. Green
- 6. Orange
- 7. Red
- 8. Pink

**Neutral/Monochrome:**
- 9. Gray
- 10. Dark Gray
- 11. Black
- 12. White

The script will generate a file named `Calendar_YYYY.pptx` in the current directory.

## Layout Specifications

- **Slide Size:** Widescreen 16:9 (13.333" × 7.5")
- **Calendar Grid:** 4 columns × 3 rows
- **Each Month:** 3.04" wide × 2" tall
- **Horizontal Gutter:** 0.2"
- **Vertical Gutter:** 0.3"
- **Font:** Arial 12pt

## Requirements

- Python 3.7+
- python-pptx

## License

MIT License - Feel free to use and modify as needed!

## Author

Created by mrfrostycakes

