🏆 Interactive Data Challenge Platform

An interactive web application built entirely on Google Apps Script, React, and Tailwind CSS. It transforms regular Google Sheets into a live, gamified challenge platform. It allows trainers to launch interactive quizzes and assessments with a live leaderboard and complete visual customization directly from the spreadsheet.

✨ Key Features

Modern & Fast UI: Built with React and Tailwind CSS to ensure a responsive experience across all smart devices and screen sizes.

Free Cloud Database: Uses Google Sheets as a Backend Database, eliminating the need for external hosting.

Live Leaderboard: Features a sidebar inside the Google Sheet for trainers to monitor participant results in real-time.

Multiple Question Types: Supports single choice, multiple choice, true/false, and ordering questions.

Dynamic Scoring System: Points are awarded based on both answer accuracy and speed. Participants are assigned custom ranks (titles) based on their performance.

Visual Identity Customization: Platform colors and themes can be fully modified directly from the sheet without touching any code.

Concurrency Safe: Utilizes LockService to ensure accurate data recording even when hundreds of users submit their answers simultaneously.

🛠️ Tech Stack

Frontend: React.js, Tailwind CSS, HTML/JS

Backend: Google Apps Script (GAS)

Database: Google Sheets

🚀 Installation and Setup

This project requires no servers or complex configurations. Follow these steps:

1. Setup Google Apps Script

Create a new Google Sheets file.

From the top menu, go to Extensions > Apps Script.

Create 3 new files in the code editor and paste the respective codes from this repository:

Code.gs (Main script file)

index.html (Trainee web app interface)

Sidebar.html (Trainer's leaderboard sidebar)

2. Initialize Database

Save all code files (Ctrl + S) and close the code editor to return to the spreadsheet.

Refresh the Google Sheets page.

A new custom menu will appear at the top named "🏆 Data Challenge".

Click on it, then select "Setup Database & Settings" (You will need to grant the required permissions during the first run).

The script will automatically generate the required sheets: Settings, Questions, and Results.

3. Deploy the Application

Return to the Apps Script editor.

At the top right, click on Deploy > New deployment.

Click the gear icon ⚙️ next to Select type and choose Web app.

Settings:

Execute as: Me

Who has access: Anyone (So participants can join)

Click Deploy and copy the resulting URL. This is the link you will share with your trainees!

📝 Content Management

Adding Questions: Open the Questions sheet to add your queries. Define the question type, options, time limit, and required points.

Customizing Colors and Ranks: Open the Settings sheet to modify the Hex color values and rank names as you desire.

Tracking Participants: From the sheet's top menu, select Show Leaderboard (Trainer) to open the live sidebar that updates automatically once the challenge begins.

🤝 Contributing

Contributions are welcome! If you have a feature request or have squashed a bug, feel free to open a Pull Request or submit an Issue.

📄 License

This project is open-source and available under the MIT License.
