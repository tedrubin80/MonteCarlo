# Monte Carlo Simulation - Railway Deployment

## Quick Start Guide

### 1. Prepare Your Files
Create a new folder and add these files:
- `main.py` (the runner script)
- `consumer_harm_monte_carlo.py` (your original simulation)
- `export_to_excel.py` (Excel export script)
- `requirements.txt` (Python packages)
- `railway.json` (Railway config)
- `Procfile` (process definition)

### 2. Deploy to Railway

#### Option A: Deploy with GitHub (Recommended)
1. Create a new GitHub repository
2. Upload all files to the repository
3. Go to [Railway Dashboard](https://railway.app/dashboard)
4. Click "New Project" → "Deploy from GitHub repo"
5. Select your repository
6. Railway will automatically start deployment

#### Option B: Deploy with Railway CLI
1. Install Railway CLI:
   ```bash
   npm install -g @railway/cli
   ```

2. Login to Railway:
   ```bash
   railway login
   ```

3. Initialize and deploy:
   ```bash
   railway init
   railway up
   ```

### 3. Monitor Deployment
- Go to your Railway dashboard
- Click on your project
- View logs to see simulation progress
- Deployment takes 3-5 minutes
- Simulation runs for 5-10 minutes

### 4. Download Results
Once complete, Railway will show:
```
SIMULATION COMPLETE!
Download simulation_results.zip to get all your files!
```

To download:
1. Go to the deployment logs
2. Look for the "Deployments" tab
3. Download the `simulation_results.zip` file
4. Extract to get:
   - `Consumer_Harm_Analysis.xlsx`
   - `monte_carlo_raw_data.csv`
   - `consumer_harm_analysis.png`
   - `simulation_summary.txt`

## What Each File Does

- **main.py**: Orchestrates the entire simulation process
- **requirements.txt**: Tells Railway which Python packages to install
- **railway.json**: Configures how Railway runs your app
- **Procfile**: Defines the command to start your simulation

## Troubleshooting

### If deployment fails:
1. Check the build logs for errors
2. Ensure all Python files are uploaded
3. Verify file names match exactly

### If simulation fails:
1. Check runtime logs
2. May need to reduce N_SIMULATIONS if memory issues
3. Contact Railway support (very responsive)

## Cost
- Free tier: $5/month credit
- This simulation: ~$0.10
- No credit card needed for free tier

## Alternative: Local Deployment
If you prefer to try fixing local issues:
```bash
# Create virtual environment
python3 -m venv monte_env
source monte_env/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run simulation
python main.py
```

## Support
- Railway Docs: https://docs.railway.app
- Railway Discord: Very helpful community
- Your simulation issues: Check logs in Railway dashboard