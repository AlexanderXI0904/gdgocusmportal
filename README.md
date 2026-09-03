# GDGoC USM Portal

A centralized internal workspace and automation portal developed for **Google Developer Group On Campus, Universiti Sains Malaysia (GDGoC USM)**. The platform hosts scalable internal tools, featuring a robust **Mass Email Application** as its flagship service.

---

## Key Features

* **Mass Email Automation:** Streamlines personalized bulk outreach by mapping data columns directly from uploaded `.xlsx` files and dispatching via the **Microsoft Graph API**.
* **Dual-Mode Editor:** Seamlessly toggle between a rich visual WYSIWYG workspace and an advanced HTML editor powered by **CodeMirror**.
* **Dynamic Template Management:** Save, export, and import campaign templates (`.json`) for seamless reuse across committee members.
* **Synchronized Dark/Light Theme:** Built-in theme engine with synchronized CSS variable transitions and custom styling matching Google design standards.
* **Multi-Tab Architecture:** Clean hub-and-spoke navigation separating the landing hub, the active Email Application, and project documentation.

---

## Tech Stack

* **Backend:** Python, Flask, Openpyxl, MSAL (Microsoft Authentication Library)
* **Frontend:** HTML5, Jinja2, Vanilla JavaScript, CSS3
* **Integrations:** Microsoft Graph API, CodeMirror 5

---

## Local Setup & Installation

1. **Clone the repository:**

   ```bash
   git clone [https://github.com/AlexanderXI0904/gdgocusmportal.git](https://github.com/AlexanderXI0904/gdgocusmportal.git)
   ```

2. **Install dependencies:**


   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**

   ```bash
   python app.py
   ```

   Open your browser and navigate to `http://localhost:5000`.

---

## Developed by

**Google Developer Group On Campus, Universiti Sains Malaysia**

*Committee, Academic Session 2026/27*