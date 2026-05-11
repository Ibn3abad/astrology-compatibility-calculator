# ✨ Astrology Compatibility Calculator  
## A Fun Python Tool for Love, Planets & Excel Charts 💫

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.x-blue?style=for-the-badge&logo=python" />
  <img src="https://img.shields.io/badge/Astrology-Fun%20Project-purple?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Excel-Report-green?style=for-the-badge&logo=microsoft-excel" />
  <img src="https://img.shields.io/badge/Made%20with-Love-red?style=for-the-badge" />
</p>

<p align="center">
  <b>A personal Python project that calculates astrology-inspired compatibility scores using planetary positions, Mars, Venus, and transit tables.</b>
</p>

---

## 🌙 About the Project

This project was created as a fun and personal tool.

My wife is interested in astrology and used to calculate compatibility manually using **Mars and Venus tables**.  
Doing this every time by hand was slow, repetitive, and a little annoying.

So I wrote this Python script to automate the calculation and generate a clear Excel report with scores, dates, zodiac information, and a visual chart.

> This project is made for fun, curiosity, and learning.  
> It is not scientific advice and should not be treated as a serious prediction tool.

---

## 💖 Why I Built It

I wanted to turn a repetitive manual astrology task into a small Python automation project.

Instead of checking tables again and again, the script can:

- calculate daily astrology-inspired scores
- analyze planetary positions
- include Mars and Venus signs
- export the result into Excel
- create a chart to make everything easier to read

It was also a nice way to combine:

- Python programming
- data processing
- Excel automation
- astronomy data
- astrology-inspired logic
- and a little bit of love ❤️

---

## 🔮 What Does It Do?

The script calculates astrology-based compatibility or transit scores for a selected year.

It compares natal planetary positions with daily transit planetary positions and gives each day a score based on astrological aspects.

The final result is saved as an Excel file.

---

## ✨ Features

- 🌍 Fetches planetary data using NASA JPL Horizons through `astroquery`
- 🪐 Calculates planetary positions
- 💕 Focuses on relationship-related planets like **Venus** and **Mars**
- ☀️ Adds special weighting for important planets such as Sun, Venus, and Mars
- 📊 Calculates daily astrology-inspired scores
- 📈 Generates an Excel chart
- ♈ Tracks Mars and Venus zodiac signs
- 📅 Shows zodiac sign transitions
- 📁 Exports everything into a clean Excel report
- 🐍 Written in Python

---

## 🧠 How It Works

The script uses a birth date as a natal reference and calculates the planetary transit positions for each day of a selected year.

Then it checks the angular distance between planets and looks for common astrological aspects.

Supported aspects include:

- Conjunction
- Sextile
- Square
- Trine
- Opposition

Each aspect contributes to the final score depending on its type and weight.

---

## 🪐 Planets and Astrology Logic

The project uses planets commonly associated with astrology compatibility and relationship interpretation.

Special attention is given to:

- **Venus** – love, attraction, harmony
- **Mars** – passion, energy, drive
- **Sun** – identity and core energy

The logic is simplified and designed for fun, not for professional astrology.

---

## 📦 Requirements

Install the required Python packages:

```bash
pip install pandas numpy astroquery astropy openpyxl
```

Main libraries used:

- `pandas`
- `numpy`
- `astroquery`
- `astropy`
- `openpyxl`

---

## 🚀 How to Use

1. Clone the repository:

```bash
git clone https://github.com/Ibn3abad/astrology-compatibility-calculator.git
```

2. Go into the project folder:

```bash
cd astrology-compatibility-calculator
```

3. Open the Python file and adjust the configuration values:

```python
GEBURTSDATUM = "1979-06-04 12:00"
JAHR = 1993
```

- `GEBURTSDATUM` is the birth date and time used as the natal reference.
- `JAHR` is the year that should be analyzed.

4. Run the script:

```bash
python astrology_compatibility_calculator.py
```

5. After the script finishes, it creates an Excel file similar to:

```text
Astrologie_Analyse_1993_DoppelAchse.xlsx
```

---

## 📊 Example Output

The generated Excel file contains:

- date
- calculated score
- normalized score in percent
- Mars zodiac sign
- Venus zodiac sign
- zodiac transition markers
- an automatically generated chart

This makes it easy to see which days have higher or lower astrology-inspired scores.

---

## 📂 Project Structure

```text
astrology-compatibility-calculator/
│
├── astrology_compatibility_calculator.py
├── README.md
└── LICENSE
```

---

## 🛠️ Technologies Used

<p align="center">
  <img src="https://img.shields.io/badge/Python-Programming-blue?style=flat-square&logo=python" />
  <img src="https://img.shields.io/badge/Pandas-Data%20Analysis-lightgrey?style=flat-square" />
  <img src="https://img.shields.io/badge/NumPy-Numerical%20Computing-blueviolet?style=flat-square" />
  <img src="https://img.shields.io/badge/OpenPyXL-Excel%20Automation-brightgreen?style=flat-square" />
  <img src="https://img.shields.io/badge/Astroquery-Astronomy%20Data-orange?style=flat-square" />
</p>

---

## 🌟 Possible Future Improvements

Some ideas for future versions:

- add a graphical user interface
- add command-line arguments
- support multiple birth dates
- compare two people directly
- add more astrology rules
- export charts as images
- add a web version
- improve the Excel layout
- add automated tests

---

## ⚠️ Disclaimer

This project is for **entertainment and learning purposes only**.

Astrology is not a scientific method, and the calculated scores should be understood as a fun interpretation rather than a factual or predictive result.

Please use this project with curiosity, humor, and an open mind.

---

## 🤝 Contributing

Contributions are welcome.

You can help by:

- improving the code
- fixing bugs
- improving the README
- adding examples
- adding tests
- suggesting new features

Example contribution workflow:

```bash
git clone https://github.com/Ibn3abad/astrology-compatibility-calculator.git
cd astrology-compatibility-calculator
git checkout -b improve-project
```

After your changes, open a Pull Request.

---

## 📄 License

This project is released under the **CC0-1.0 License**.

You are free to use, modify, and share it.

---

## ⭐ Support the Project

If you like this project, you can support it by:

- giving the repository a Star ⭐
- sharing it with friends
- trying the code
- suggesting improvements
- building your own fun version

---

## ✍️ Author

**Obayda**

A fun Python astrology project built with love, curiosity, and a little bit of planetary magic. 🌙✨
