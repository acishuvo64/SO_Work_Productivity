ACI Premio Plastics — SO Productivity Dashboard

FILES
- dashboard.py: Streamlit dashboard application
- requirements.txt: Python packages
- data/SO_Activity_Attendance.xlsx: source workbook included in this starter package

GITHUB SETUP
1. Create a GitHub repository. Because this workbook contains employee/work performance information, use a PRIVATE repository and only grant access to authorized people.
2. Upload dashboard.py, requirements.txt, and the data folder with SO_Activity_Attendance.xlsx.
3. In Streamlit Community Cloud, create/deploy an app from the repository, set Main file path to dashboard.py.
4. Once deployment succeeds, open the generated Streamlit URL and test the filters and report totals before sharing it with the Director.
5. For monthly updates, replace data/SO_Activity_Attendance.xlsx with the latest workbook while keeping the same filename and sheet names, then commit/push. Streamlit Cloud will redeploy.

IMPORTANT DATA NOTES
- Required sheet names are exactly: Summary Report, SO Activity, Attendance (the uploaded workbook's third sheet has a trailing space in its name: 'Attendance ').
- Targets in Summary Report are used for outlet and memo achievement.
- The attendance workbook does not contain a Total Hours column; actual work hours cannot be calculated until the source includes a reliable duration field.
- SO Activity has Market Name but no individual outlet name; repeat analysis is market-level.
- The AI-style summary is rule-based and does not call an external AI service.
- Hiding code in the dashboard UI does not hide code or data from repository collaborators. A public GitHub repository can expose both code and committed workbook data. Keep the repository private for confidential employee data.
