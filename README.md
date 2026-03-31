# ⛪ SDA Kubwa Announcement System

This system automatically converts weekly church announcements from raw text into a professional, looping PowerPoint presentation.

## 📁 File Structure (Flat)

- **`run_announcements.bat`**: The only file you need to double-click. It handles everything.
- **`build_slides.py`**: The "brain" of the system (Python script).
- **`announcements.json`**: The data file where the AI output is pasted.
- **`SDA_Template.pptx`**: The master design template.
- **`SDA_Kubwa_Announcements.pptx`**: The final output file created by the script.

---

## 🚀 How to Use (Weekly Workflow)

### Step 1: Generate the Data

1.  Copy your raw announcement text (from WhatsApp, Bulletin, or Paper).
2.  Use the **AI Weekly Prompt** to process the text.
3.  The AI will give you a block of code (JSON). Copy that code.

### Step 2: Update the Data File

1.  Open `announcements.json` with **Notepad**.
2.  Delete everything inside and **Paste** the new JSON from the AI.
3.  **Save** (Ctrl + S) and close the file.

### Step 3: Create the PowerPoint

1.  Double-click **`run_announcements.bat`**.
2.  A window will appear and show progress:
    - _Checking dependencies..._
    - _Generating PowerPoint..._
    - _Opening newest presentation..._
3.  The PowerPoint will open automatically. Press **F5** to start the slideshow.

---

## 🛠️ Troubleshooting

| Issue                  | Solution                                                                                                                                    |
| :--------------------- | :------------------------------------------------------------------------------------------------------------------------------------------ |
| **Window turns RED**   | There is an error. Read the message in the window. Usually, it means `announcements.json` has a typo (like a missing comma).                |
| **"Python not found"** | Ensure Python is installed from [python.org](https://www.python.org/) and the box **"Add Python to PATH"** was checked during installation. |
| **Permission Denied**  | Close the PowerPoint file if it is already open and run the `.bat` file again.                                                              |
| **Design looks wrong** | Edit the `SDA_Template.pptx` file to change colors, fonts, or logos. The script will follow your changes.                                   |

---

## 📝 Maintenance Notes

- **Slide Timing:** Slides are set to advance automatically every **8 seconds**.
- **Looping:** The presentation is hard-coded to loop continuously until "Esc" is pressed.
- **Icons:** Use standard emojis in the JSON for the "icon" field; the script will render them on the top-right of each slide.

---

> **Tip:** To make this even easier, right-click `run_announcements.bat` and select **Send to > Desktop (create shortcut)**. You can rename the shortcut to "Generate Church Slides."
