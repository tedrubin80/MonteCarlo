#!/usr/bin/env python3
"""
Standalone Monte Carlo Simulation - All functionality in one file
This combines the simulation, Excel export, and main runner
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import os
import sys
import json
import zipfile
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Set random seed for reproducibility
np.random.seed(42)

# Configuration
N_SIMULATIONS = 10000  # Number of simulated customers
ANNUAL_TRANSACTIONS = 1.73e6  # 1.73 million transactions per year
SERVICE_FAILURE_PENALTY = 1000

# Distribution parameters (using triangular distributions)
PARAMS = {
    'base_service_cost': {'min': 2500, 'mode': 3200, 'max': 4000},
    'hidden_fees': {'min': 0, 'mode': 375, 'max': 1100},
    'service_failure_prob': {'min': 0.15, 'mode': 0.30, 'max': 0.45},
    'claim_denial_prob': {'min': 0.60, 'mode': 0.85, 'max': 0.95},
    'damage_occurrence_rate': {'min': 0.05, 'mode': 0.12, 'max': 0.25},
    'average_damage_value': {'min': 500, 'mode': 2500, 'max': 10000}
}

def triangular_sample(min_val, mode_val, max_val, size):
    """Generate samples from triangular distribution"""
    return np.random.triangular(min_val, mode_val, max_val, size)

def run_monte_carlo_simulation(params=PARAMS, n_sims=N_SIMULATIONS):
    """Run Monte Carlo simulation for consumer harm"""
    # Generate random samples for each parameter
    service_costs = triangular_sample(
        params['base_service_cost']['min'],
        params['base_service_cost']['mode'],
        params['base_service_cost']['max'],
        n_sims
    )
    
    hidden_fees = triangular_sample(
        params['hidden_fees']['min'],
        params['hidden_fees']['mode'],
        params['hidden_fees']['max'],
        n_sims
    )
    
    service_failure_probs = triangular_sample(
        params['service_failure_prob']['min'],
        params['service_failure_prob']['mode'],
        params['service_failure_prob']['max'],
        n_sims
    )
    
    claim_denial_probs = triangular_sample(
        params['claim_denial_prob']['min'],
        params['claim_denial_prob']['mode'],
        params['claim_denial_prob']['max'],
        n_sims
    )
    
    damage_occurrence_rates = triangular_sample(
        params['damage_occurrence_rate']['min'],
        params['damage_occurrence_rate']['mode'],
        params['damage_occurrence_rate']['max'],
        n_sims
    )
    
    damage_values = triangular_sample(
        params['average_damage_value']['min'],
        params['average_damage_value']['mode'],
        params['average_damage_value']['max'],
        n_sims
    )
    
    # Simulate events
    service_failures = np.random.random(n_sims) < service_failure_probs
    damage_occurred = np.random.random(n_sims) < damage_occurrence_rates
    claims_denied = np.random.random(n_sims) < claim_denial_probs
    
    # Calculate harm components
    service_failure_harm = service_failures * SERVICE_FAILURE_PENALTY
    damage_harm = damage_occurred * damage_values * claims_denied
    
    # Total harm per customer
    total_harm = hidden_fees + service_failure_harm + damage_harm
    
    # Create results dataframe
    results = pd.DataFrame({
        'service_cost': service_costs,
        'hidden_fees': hidden_fees,
        'service_failure': service_failures,
        'service_failure_harm': service_failure_harm,
        'damage_occurred': damage_occurred,
        'damage_value': damage_values,
        'claim_denied': claims_denied,
        'damage_harm': damage_harm,
        'total_harm': total_harm
    })
    
    return results

def calculate_statistics(results):
    """Calculate key statistics from simulation results"""
    harm = results['total_harm']
    
    stats_dict = {
        'Mean Harm': harm.mean(),
        'Median Harm': harm.median(),
        'Std Dev': harm.std(),
        'Min Harm': harm.min(),
        'Max Harm': harm.max(),
        '10th Percentile': harm.quantile(0.10),
        '25th Percentile': harm.quantile(0.25),
        '75th Percentile': harm.quantile(0.75),
        '90th Percentile': harm.quantile(0.90),
        '95th Percentile': harm.quantile(0.95),
        '99th Percentile': harm.quantile(0.99),
        'Customers with Zero Harm': (harm == 0).sum(),
        'Customers with Harm > $1000': (harm > 1000).sum(),
        'Customers with Harm > $5000': (harm > 5000).sum(),
        '% with Zero Harm': (harm == 0).sum() / len(harm) * 100,
        '% with Harm > $1000': (harm > 1000).sum() / len(harm) * 100,
        '% with Harm > $5000': (harm > 5000).sum() / len(harm) * 100,
        'Annual Industry Impact (Mean)': harm.mean() * ANNUAL_TRANSACTIONS,
        'Annual Industry Impact (95th %ile)': harm.quantile(0.95) * ANNUAL_TRANSACTIONS
    }
    
    return stats_dict

def create_excel_report(results, stats, scenario_results, output_dir):
    """Create Excel report with basic pandas functionality"""
    excel_file = os.path.join(output_dir, "Consumer_Harm_Analysis.xlsx")
    
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        # Executive Summary
        summary_data = {
            'Metric': [
                'Mean Consumer Harm',
                'Median Consumer Harm',
                '95th Percentile Harm',
                'Annual Industry Impact (Mean)',
                'Customers with Zero Harm',
                'Customers with Harm > $1,000',
                'Customers with Harm > $5,000'
            ],
            'Value': [
                f"${stats['Mean Harm']:,.2f}",
                f"${stats['Median Harm']:,.2f}",
                f"${stats['95th Percentile']:,.2f}",
                f"${stats['Annual Industry Impact (Mean)']:,.0f}",
                f"{stats['Customers with Zero Harm']:,}",
                f"{stats['Customers with Harm > $1000']:,}",
                f"{stats['Customers with Harm > $5000']:,}"
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
        
        # Detailed Results (first 1000)
        results.head(1000).to_excel(writer, sheet_name='Detailed Results', index=False)
        
        # Percentile Analysis
        percentiles = [1, 5, 10, 25, 50, 75, 90, 95, 99]
        percentile_data = {
            'Percentile': [f"{p}th" for p in percentiles],
            'Harm Amount': [results['total_harm'].quantile(p/100) for p in percentiles],
            '% of Customers': percentiles
        }
        percentile_df = pd.DataFrame(percentile_data)
        percentile_df.to_excel(writer, sheet_name='Percentile Analysis', index=False)
        
        # Harm Components
        components_data = {
            'Component': ['Hidden Fees', 'Service Failures', 'Damages (Denied Claims)'],
            'Mean': [
                results['hidden_fees'].mean(),
                results['service_failure_harm'].mean(),
                results['damage_harm'].mean()
            ],
            'Median': [
                results['hidden_fees'].median(),
                results['service_failure_harm'].median(),
                results['damage_harm'].median()
            ],
            'Max': [
                results['hidden_fees'].max(),
                results['service_failure_harm'].max(),
                results['damage_harm'].max()
            ]
        }
        components_df = pd.DataFrame(components_data)
        components_df.to_excel(writer, sheet_name='Harm Components', index=False)
        
        # Scenario Comparison
        scenario_data = []
        for scenario_name, scenario_info in scenario_results.items():
            scenario_data.append({
                'Scenario': scenario_name,
                'Mean Harm': scenario_info['stats']['Mean Harm'],
                'Median Harm': scenario_info['stats']['Median Harm'],
                '95th Percentile': scenario_info['stats']['95th Percentile'],
                'Annual Impact': scenario_info['stats']['Annual Industry Impact (Mean)']
            })
        scenario_df = pd.DataFrame(scenario_data)
        scenario_df.to_excel(writer, sheet_name='Scenario Comparison', index=False)
    
    return excel_file

def create_output_directory():
    """Create output directory for results"""
    # Use /data if available, otherwise current directory
    if os.path.exists("/data"):
        output_dir = "/data/simulation_results"
        print("Using persistent storage at /data")
    else:
        output_dir = "simulation_results"
        print("Using local storage")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    return output_dir

def main():
    """Main execution function"""
    print("="*60)
    print("CONSUMER HARM MONTE CARLO SIMULATION")
    print("Standalone Version - Railway Deployment")
    print("="*60)
    print()
    
    # Create output directory
    output_dir = create_output_directory()
    
    # Run base simulation
    print(f"Step 1: Running base simulation with {N_SIMULATIONS:,} iterations...")
    results = run_monte_carlo_simulation()
    stats = calculate_statistics(results)
    
    print("✓ Base simulation complete")
    print(f"  - Mean harm: ${stats['Mean Harm']:,.2f}")
    print(f"  - 95th percentile: ${stats['95th Percentile']:,.2f}")
    print(f"  - Annual impact: ${stats['Annual Industry Impact (Mean)']:,.0f}")
    
    # Run scenario analysis
    print("\nStep 2: Running scenario analysis...")
    scenarios = {
        'Status Quo': PARAMS,
        'Moderate Reform': {
            'base_service_cost': PARAMS['base_service_cost'],
            'hidden_fees': {'min': 0, 'mode': 150, 'max': 500},
            'service_failure_prob': {'min': 0.10, 'mode': 0.20, 'max': 0.30},
            'claim_denial_prob': {'min': 0.40, 'mode': 0.60, 'max': 0.80},
            'damage_occurrence_rate': PARAMS['damage_occurrence_rate'],
            'average_damage_value': PARAMS['average_damage_value']
        },
        'Strong Reform': {
            'base_service_cost': PARAMS['base_service_cost'],
            'hidden_fees': {'min': 0, 'mode': 50, 'max': 200},
            'service_failure_prob': {'min': 0.05, 'mode': 0.10, 'max': 0.15},
            'claim_denial_prob': {'min': 0.20, 'mode': 0.35, 'max': 0.50},
            'damage_occurrence_rate': {'min': 0.03, 'mode': 0.08, 'max': 0.15},
            'average_damage_value': PARAMS['average_damage_value']
        }
    }
    
    scenario_results = {}
    for scenario_name, scenario_params in scenarios.items():
        print(f"  - Running {scenario_name} scenario...")
        scenario_res = run_monte_carlo_simulation(scenario_params, n_sims=N_SIMULATIONS)
        scenario_stats = calculate_statistics(scenario_res)
        scenario_results[scenario_name] = {
            'results': scenario_res,
            'stats': scenario_stats
        }
    
    print("✓ Scenario analysis complete")
    
    # Generate visualizations
    print("\nStep 3: Generating visualizations...")
    try:
        import matplotlib
        matplotlib.use('Agg')  # Use non-interactive backend
        
        # Create summary visualization
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 10))
        
        # 1. Histogram of total harm
        ax1.hist(results['total_harm'], bins=50, color='steelblue', alpha=0.7, edgecolor='black')
        ax1.axvline(results['total_harm'].mean(), color='red', linestyle='--', linewidth=2, 
                    label=f'Mean: ${results["total_harm"].mean():.0f}')
        ax1.axvline(results['total_harm'].median(), color='green', linestyle='--', linewidth=2, 
                    label=f'Median: ${results["total_harm"].median():.0f}')
        ax1.set_xlabel('Total Consumer Harm ($)')
        ax1.set_ylabel('Frequency')
        ax1.set_title('Distribution of Consumer Harm')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        
        # 2. Scenario comparison
        scenarios_list = list(scenarios.keys())
        mean_harms = [scenario_results[s]['stats']['Mean Harm'] for s in scenarios_list]
        colors = ['red', 'orange', 'green']
        
        bars = ax2.bar(scenarios_list, mean_harms, color=colors)
        ax2.set_ylabel('Mean Harm ($)')
        ax2.set_title('Average Consumer Harm by Scenario')
        for bar, harm in zip(bars, mean_harms):
            ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 10,
                     f'${harm:.0f}', ha='center')
        
        # 3. Percentile chart
        percentiles = [10, 25, 50, 75, 90, 95, 99]
        percentile_values = [results['total_harm'].quantile(p/100) for p in percentiles]
        
        ax3.bar([str(p) + 'th' for p in percentiles], percentile_values, color='coral')
        ax3.set_xlabel('Percentile')
        ax3.set_ylabel('Harm Amount ($)')
        ax3.set_title('Consumer Harm by Percentile')
        for i, v in enumerate(percentile_values):
            ax3.text(i, v + 50, f'${v:.0f}', ha='center', va='bottom', fontsize=8)
        
        # 4. Component breakdown
        component_means = [
            results['hidden_fees'].mean(),
            results['service_failure_harm'].mean(),
            results['damage_harm'].mean()
        ]
        labels = ['Hidden Fees', 'Service Failures', 'Damages (Denied)']
        colors_pie = ['#ff9999', '#66b3ff', '#99ff99']
        
        ax4.pie(component_means, labels=labels, colors=colors_pie, autopct='%1.1f%%', startangle=90)
        ax4.set_title('Average Harm Breakdown by Component')
        
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, 'consumer_harm_analysis.png'), dpi=300, bbox_inches='tight')
        plt.close()
        
        print("✓ Visualizations created")
    except Exception as e:
        print(f"⚠ Visualization error (non-critical): {str(e)}")
    
    # Generate Excel report
    print("\nStep 4: Generating Excel report...")
    try:
        excel_file = create_excel_report(results, stats, scenario_results, output_dir)
        print("✓ Excel report generated")
    except Exception as e:
        print(f"⚠ Excel generation error: {str(e)}")
        excel_file = None
    
    # Save raw data
    print("\nStep 5: Saving raw data...")
    results.to_csv(os.path.join(output_dir, 'monte_carlo_raw_data.csv'), index=False)
    
    # Create summary report
    summary_file = os.path.join(output_dir, 'simulation_summary.txt')
    with open(summary_file, 'w') as f:
        f.write("CONSUMER HARM MONTE CARLO SIMULATION SUMMARY\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"Simulation Date: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}\n")
        f.write(f"Platform: Railway Cloud (Standalone Version)\n")
        f.write(f"Number of simulations: {N_SIMULATIONS:,}\n")
        f.write(f"Annual transactions: {ANNUAL_TRANSACTIONS:,.0f}\n\n")
        
        for scenario_name, scenario_data in scenario_results.items():
            f.write(f"\n{scenario_name.upper()} SCENARIO\n")
            f.write("-" * 40 + "\n")
            s = scenario_data['stats']
            f.write(f"Mean Harm: ${s['Mean Harm']:,.2f}\n")
            f.write(f"Median Harm: ${s['Median Harm']:,.2f}\n")
            f.write(f"95th Percentile: ${s['95th Percentile']:,.2f}\n")
            f.write(f"99th Percentile: ${s['99th Percentile']:,.2f}\n")
            f.write(f"Annual Industry Impact: ${s['Annual Industry Impact (Mean)']:,.0f}\n")
            
            if scenario_name != 'Status Quo':
                reduction = (1 - s['Mean Harm'] / scenario_results['Status Quo']['stats']['Mean Harm']) * 100
                f.write(f"Reduction from Status Quo: {reduction:.1f}%\n")
    
    print("✓ Summary report created")
    
    # If using /data, copy key files to root for easy access
    if os.path.exists("/data"):
        import shutil
        key_files = [
            'Consumer_Harm_Analysis.xlsx',
            'monte_carlo_raw_data.csv',
            'simulation_summary.txt',
            'consumer_harm_analysis.png'
        ]
        
        print("\nCopying key files to /data root for easy access...")
        for filename in key_files:
            src_path = os.path.join(output_dir, filename)
            if os.path.exists(src_path):
                dst_path = os.path.join('/data', filename)
                shutil.copy2(src_path, dst_path)
                print(f"✓ Copied {filename} to /data/")
    
    print("\n" + "="*60)
    print("SIMULATION COMPLETE!")
    print("="*60)
    
    if os.path.exists("/data"):
        print("\nFiles saved to persistent storage:")
        print(f"  ✓ Results directory: {output_dir}")
        print("\nKey files also available at:")
        print("  ✓ /data/Consumer_Harm_Analysis.xlsx")
        print("  ✓ /data/monte_carlo_raw_data.csv")
        print("  ✓ /data/simulation_summary.txt")
        print("  ✓ /data/consumer_harm_analysis.png")
    else:
        print("\nGenerated files:")
        print("  ✓ Consumer_Harm_Analysis.xlsx - Excel report")
        print("  ✓ monte_carlo_raw_data.csv - Raw simulation data")
        print("  ✓ consumer_harm_analysis.png - Visualizations")
        print("  ✓ simulation_summary.txt - Key findings")
    
    print("\n✅ Success! The simulation completed successfully.")
    return True

if __name__ == "__main__":
    try:
        success = main()
        if success:
            sys.exit(0)
        else:
            sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)