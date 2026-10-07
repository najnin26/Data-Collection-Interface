# Bangla Pragmatics Data Collection — Streamlit

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

The app contains 25 Bangla pragmatic scenarios focused on university student life
and stores each completed response in `data/bangla_pragmatics_responses.xlsx`.

## Download collected data
In Streamlit Cloud, open the app settings and add an `admin_password` value under
Secrets, for example:

```toml
admin_password = "use-a-long-unique-password"
```

Do not commit this password to the repository. In the app, open **Researcher data
export**, enter the password, and choose **Download response workbook**. The
download is a copy of the workbook on the Streamlit server; it does not update the
workbook in your local project folder. Download backups regularly because
Streamlit Cloud's local filesystem may not persist across restarts or redeploys.

## Excel fields
- Participant demographics
- Scenario ID/category
- Context
- Participant's natural Bangla utterance
- Participant's intended meaning

## Research note
Before recruiting participants, obtain any required university ethics/IRB approval.
Do not collect unnecessary identifying information.
