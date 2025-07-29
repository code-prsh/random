# Cold Email Launcher

A Streamlit-based web application for sending personalized cold emails to multiple recipients with batch processing.

## Features
- Upload Excel/CSV files with contact information
- Customizable email templates
- Batch processing with configurable batch sizes
- Real-time progress tracking
- SMTP integration for sending emails

## Deployment Options

### Option 1: Streamlit Sharing (Recommended)
1. Push this repository to GitHub
2. Sign up at [Streamlit Sharing](https://share.streamlit.io/)
3. Click "New App" and connect your GitHub repository
4. Select the main branch and set the main file to `app.py`
5. Click "Deploy!"

### Option 2: Heroku
1. Install the [Heroku CLI](https://devcenter.heroku.com/articles/heroku-cli)
2. Run the following commands:
   ```bash
   heroku login
   heroku create your-app-name
   git add .
   git commit -m "Initial commit"
   git push heroku main
   heroku ps:scale web=1
   ```

### Option 3: Local Development
1. Clone the repository
2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the app:
   ```bash
   streamlit run app.py
   ```

## Environment Variables
Create a `.env` file in the root directory with your SMTP settings:
```
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
SMTP_USERNAME=your-email@gmail.com
SMTP_PASSWORD=your-app-password
```

## Security Note
- Never commit sensitive information like email credentials
- Use environment variables for configuration
- Ensure your SMTP provider allows programmatic access

## License
MIT License
