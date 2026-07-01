# 🌬️ Wind Rose Generator

A professional, customizable **Wind Rose Diagram Generator** built with Streamlit for environmental engineers, EIA consultants, researchers, and wind resource analysts.

This tool was developed to provide high-quality, publication-ready wind rose diagrams and supporting data tables with extensive customization options — matching or exceeding the output quality of commercial tools like WindRose PRO, while remaining completely free and flexible.

---

## ✨ Key Features

### Wind Rose Visualizations
- **Interactive Plotly Wind Rose** — Frequency-based with color by average wind speed
- **Classic Layered Wind Rose** (Report Style) — Stacked by wind speed classes with full customization
- **Average Wind Speed Rose** — Polar plot showing mean wind speed per direction

### Powerful Customization
- Multiple color palettes (Image Style, Viridis, Plasma, RdYlBu, Coolwarm)
- Toggle between **Percentage of time** and **Number of hourly events**
- Reverse stacking order (strongest winds inside or outside)
- Show/hide direction labels, value labels, grid lines
- Toggle calm circle in center with percentage
- Custom title support
- Adjustable PNG export resolution (150 / 300 / 600 DPI)

### Data & Reporting
- Professional **Excel export** (Enviroware / WindRose PRO style format)
- High-resolution PNG downloads for all visualizations
- Clean frequency table (direction × wind speed class matrix)
- Summary statistics and speed class distribution

### Technical
- Fetches free global historical wind data from **Open-Meteo API** (no API key required)
- Persistent state — diagrams and customizations survive widget interactions
- Clean, modern interface suitable for professional reports and theses

---

## 🚀 Getting Started

### 1. Installation

```bash
git clone <your-repo-url>
cd wind-rose-generator

pip install -r requirements.txt
```

### 2. Run the App

```bash
streamlit run wind_rose_app.py
```

The app will open in your default browser.

---

## 📖 How to Use

1. **Select Location**
   - Use the quick-select presets (includes Udupi/Mangalore area) or enter custom latitude/longitude.

2. **Set Time Period**
   - Choose start and end dates (default = last 1 year).

3. **Click "Fetch Data & Generate Wind Rose"**

4. **Customize**
   - Expand **⚙️ Customize Layered Wind Rose** to change colors, toggle labels, add custom titles, adjust DPI, enable calm circle, reverse stacking, etc.

5. **Export**
   - Download high-resolution PNGs
   - Export professional Excel report (ready for EIA submissions)
   - (Future) One-click PDF multi-page report

---

## 🛠️ Customization Options

| Option                    | Description                                      | Default     |
|---------------------------|--------------------------------------------------|-------------|
| Color Palette             | Image Style, Viridis, Plasma, RdYlBu, Coolwarm   | Image Style |
| Display Mode              | Percentage of time / Number of events            | Percentage  |
| Reverse Stacking          | Strongest winds inside vs outside                | Off         |
| Show Direction Letters    | Toggle N, NNE, NE... labels                      | On          |
| Show Value Labels         | Toggle % or count numbers outside petals         | On          |
| Show Grid Lines           | Background grid                                  | On          |
| Show Calm Circle          | Center circle with calm %                        | On          |
| Custom Title              | Add project/site name to the figure              | None        |
| DPI                       | PNG export resolution                            | 300         |

---

## 📁 Project Structure

```
wind-rose-generator/
├── wind_rose_app.py          # Main Streamlit application
├── requirements.txt          # Python dependencies
├── README.md                 # This file
└── artifacts/                # Generated outputs (Excel, PNGs)
```

---

## 📊 Data Source

- **Open-Meteo Historical Weather API** (https://open-meteo.com)
- Free, no API key required
- Global coverage with good accuracy for most locations
- Wind speed in km/h (converted to m/s for reporting)

> **Note**: This tool is intended for research, EIA, and educational purposes. For critical operational decisions, validate against local meteorological station data.

---

## 🗺️ Roadmap / Future Improvements

- [ ] One-click **PDF Report Export** (multi-page professional report)
- [ ] **Upload your own data** (CSV/Excel support)
- [ ] **Seasonal / Monthly comparison** mode
- [ ] Weibull distribution fitting for wind speed
- [ ] Wind power density / energy rose
- [ ] Interactive map for location selection
- [ ] Batch processing for multiple sites

---

## 🙏 Credits & Acknowledgments

This application was built as a specialized tool for environmental engineering workflows, particularly for researchers working on air quality, EIA, and wind resource assessment in coastal regions.

Special thanks to the user for detailed feedback during development, which helped shape many of the professional features.

---

## 📝 License

This project is provided as-is for research and professional use. Feel free to modify and extend it for your own needs.

---

**Made with care for environmental professionals and researchers.**

If you find this tool useful in your work (thesis, EIA reports, research papers), I’d love to hear about it! 

Your feedback and suggestions for improvement are always welcome.# windrose
You can make wind rose diagrams using this
