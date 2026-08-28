# GDGoC Microsoft Graph Email Sender

A robust, professional-grade Flask-based bulk email sender designed for local and team use. This tool integrates the Microsoft Graph API and Microsoft Entra ID using MSAL for secure email dispatching without requiring SMTP or app passwords. It processes Excel data to create highly personalized mass mailing campaigns. 

Designed for Google Developer Groups on Campus (GDGoC) Universiti Sains Malaysia (USM), it features a cohesive UI, dynamic live previews, campaign state persistence, and full dark mode support.

---
## 🌟 Features
* **Microsoft Authentication:** Secure, device-code flow authentication using Microsoft Entra ID and MSAL. Requires only your Client ID, Tenant ID, and the Microsoft Graph `Mail.Send` delegated permission.
* **Dynamic Excel Data Mapping:** Upload `.xlsx` or `.xlsm` files (which are stored locally inside the `uploads/` directory). You can seamlessly use dynamic Excel column insertion (e.g., `{{CompanyName}}`) directly into your templates.
* **Advanced Email Configuration:** Full CC management and native support for Microsoft Graph importance settings.
* **Dual Editors:** Compose messages using the visual WYSIWYG email editor or dive into the code with the HTML source editor.
* **Live Email Preview:** Cycle through individual recipient data on the fly with a live email preview to verify parsed variables before sending.
* **Campaign Persistence:** Export and import campaign designs (HTML, subject, CC lists) as lightweight JSON files.
* **Automated Security:** Employs an `atexit` cleanup routine to ensure no residual Excel data remains on the server upon application exit.
* **Modern, Responsive UI:** Features a GDGoC-branded design, smooth scroll transitions, and dynamic Dark/Light mode toggling.

---
## 🚀 Setup & Installation

### 1. Local Environment Setup
Ensure you have Python 3.8+ installed on your machine.

```bash
# Clone the repository
git clone [https://github.com/YOUR-USERNAME/gdgoc-email-sender.git](https://github.com/YOUR-USERNAME/gdgoc-email-sender.git)
cd gdgoc-email-sender

# Install the required Python packages
pip install flask msal openpyxl requests

# Run the application
python app.py
```
The app will be available locally at http://127.0.0.1:5000.

---
### 2. Microsoft Entra Admin Center Setup (Required)
To send emails, the application requires a **Client ID** and **Tenant ID** registered through Microsoft. Follow these steps to generate them:

1. Log in to the [Microsoft Entra Admin Center](https://entra.microsoft.com/) using your Microsoft account.
2. In the left sidebar, navigate to **Identity** > **Applications** > **App registrations**.
3. Click **New registration** at the top.
4. Name your application (e.g., `GDGoC Email Sender`).
5. Under **Supported account types**, select **Accounts in any organizational directory and personal Microsoft accounts** (to allow both school/work and personal emails).
6. Click **Register**.
7. You will be redirected to the app's Overview page. Copy and save the **Application (client) ID** and **Directory (tenant) ID**.
8. **Enable Device Code Flow:** On the left menu, go to **Authentication**. Scroll down to **Advanced settings** and toggle *Allow public client flows* to **Yes**. Click **Save** at the bottom.
9. **Grant API Permissions:** Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**. Search for `Mail.Send`, check the box, and click **Add permissions**.

Enter the copied Client ID and Tenant ID into the web application's configuration page to authenticate.

---
## 📖 Usage Guide

1. **Authenticate:** Enter your Microsoft IDs and follow the prompt to link your account via the device code flow.
2. **Upload Data:** Upload an Excel file containing your recipient data. Ensure the email addresses are in a single column.
3. **Configure Campaign:** Write your email subject and body. Use the exact column headers wrapped in double curly braces (e.g., `{{FirstName}}`) to insert dynamic variables.
4. **Preview & Send:** Review the live data preview to ensure variables are replacing correctly, then confirm the dispatch. 
5. **Save State:** Use the "Save Campaign" button to export your design to a `.json` file for future use.

---
## 🤝 Contributing

Contributions are welcome! If you are new to GitHub or open-source collaboration, we highly recommend checking out our beginner-friendly Git tutorial repository first:
👉 **[Time-to-Git-Good Guide](https://github.com/AlexanderXI0904/Time-to-Git-Good)**

**To contribute to this project:**

1. **Fork the Repository:** Click the "Fork" button at the top right of this page to create a copy in your own GitHub account.
2. **Clone your Fork:**
   ```bash
   git clone [https://github.com/YOUR-USERNAME/gdgoc-email-sender.git](https://github.com/YOUR-USERNAME/gdgoc-email-sender.git)
   cd gdgoc-email-sender
   ```

3. **Create a Branch:**
    ```bash
    git checkout -b feature/your-amazing-feature
    ```

4. **Make Changes & Commit:**
    ```bash
    git add .
    git commit -m "Add amazing feature"
    ```

5. **Push to your Fork:**
    ```bash
    git push origin feature/your-amazing-feature
    ```
Open a Pull Request: Navigate back to the original repository on GitHub and click "Compare & pull request" to submit your changes for review.