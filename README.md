# 🤖 Google Gemini AI in Excel VBA

Use Google's Gemini AI directly in Excel! Analyze data, generate formulas, summarize text, and more - all without leaving your spreadsheet.

## 🚀 Quick Start

### Step 1: Get Your API Key
1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. Sign in with your Google account
3. Click **"Create API Key"**
4. Copy your API key

### Step 2: Add the VBA Modules to Excel
1. Open your Excel workbook
2. Press `Alt + F11` to open the VBA Editor
3. Right-click on **VBAProject** → **Import File...**
4. Import these files from the `src` folder (in this order):
   - `Dictionary.cls`
   - `JsonConverter.bas`
   - `mGemini.bas`
   - `mGeminiDemo.bas` *(optional - creates demo examples)*

### Step 3: Add Your API Key
1. In the VBA Editor, open the `mGemini` module
2. Find this line near the top:
   ```vba
   Const GEMINI_API_KEY As String = "YOUR_API_KEY"
   ```
3. Replace `YOUR_API_KEY` with your actual API key

### Step 4: Run Gemini!
- **Method 1 - Macro:** Press `Alt + F8` → Select `Gemini` → Click **Run**
- **Method 2 - Formula:** Use `=AskGemini("Your question here")` in any cell

## 📖 Usage Examples

| Select Data | Run Gemini | Type Instruction |
|-------------|------------|------------------|
| Sales table | `Alt+F8` → Gemini | "Analyze trends" |
| Customer feedback | `Alt+F8` → Gemini | "Summarize sentiment" |
| Email list | `Alt+F8` → Gemini | "Extract first names" |
| *(no selection)* | `Alt+F8` → Gemini | "Write a VLOOKUP example" |

## ⚙️ Configuration Options

Edit these constants in `mGemini.bas` to customize behavior:

| Setting | Options | Description |
|---------|---------|-------------|
| `GEMINI_MODEL` | `gemini-2.5-flash`, `gemini-2.5-flash-lite`, `gemini-3-pro-preview` | AI model to use |
| `GEMINI_INPUT_MODE` | `both`, `selection`, `inputbox`, `auto` | How prompts are collected |
| `GEMINI_OUTPUT_MODE` | `lines`, `single` | How responses are displayed |

## Video Tutorial
[![YouTube Video](https://img.youtube.com/vi/_107AmTE21c/0.jpg)](https://youtu.be/_107AmTE21c)


## 💡 Explore My Solutions
From Excel add-ins to messaging tools and logistics solutions.

👉 **[See All Solutions](https://pythonandvba.com/solutions)**## 🤝 Connect with Me
- 📺 **YouTube:** [CodingIsFun](https://youtube.com/c/CodingIsFun)
- 🌐 **Website:** [PythonAndVBA](https://pythonandvba.com)
- 💬 **Discord:** [Join our Community](https://pythonandvba.com/discord)
- 💼 **LinkedIn:** [Sven Bosau](https://www.linkedin.com/in/sven-bosau/)
- 📸 **Instagram:** [Follow me](https://www.instagram.com/sven_bosau/)

## ☕️ Support My Work
Love my content and want to show appreciation? Why not [buy me a coffee](https://pythonandvba.com/coffee-donation) to fuel my creative engine? Your support means the world to me! 😊

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://pythonandvba.com/coffee-donation)

## 💌 Feedback
Got some thoughts or suggestions? Don't hesitate to reach out to me at contact@pythonandvba.com. I'd love to hear from you! 💡
![Logo](https://www.pythonandvba.com/banner-img)
