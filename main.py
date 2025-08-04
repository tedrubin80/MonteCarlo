#!/usr/bin/env python3
"""
Monte Carlo Simulation Runner for Railway
This script orchestrates the simulation and generates all outputs
"""

import os
import sys
import json
import zipfile
from datetime import datetime

# Import the simulation modules
from consumer_harm_monte_carlo import (
    run_monte_carlo_simulation,
    calculate_statistics,
    PARAMS,
    ANNUAL_TRANSACTIONS,
    N_SIMULATIONS
)
from export_to_excel import ExcelExporter

def create_output_directory():
    """Create output directory for results"""
    output_dir = "simulation_results"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    return output_dir

def run_complete_simulation():
    """Run the complete Monte Carlo simulation with all outputs"""
    print("="*60)
    print("CONSUMER HARM MONTE CARLO SIMULATION")
    print("Running on Railway Cloud Platform")
    print("="*60)
    print()
    
    # Create output directory
    output_dir = create_output_directory()
    
    # Run base simulation
    print("Step 1: Running base simulation with {:,} iterations...".format(N_SIMULATIONS))
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
        import matplotlib.pyplot as plt
        
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
    excel_file = os.path.join(output_dir, "Consumer_Harm_Analysis.xlsx")
    
    exporter = ExcelExporter(excel_file)
    exporter.create_summary_sheet(stats, scenario_results)
    exporter.create_detailed_results_sheet(results)
    exporter.create_percentile_analysis_sheet(results)
    exporter.create_harm_components_sheet(results)
    exporter.create_scenario_comparison_sheet(scenario_results)
    exporter.create_charts_sheet(results, scenario_results)
    exporter.create_parameters_sheet()
    exporter.save_workbook()
    
    print("✓ Excel report generated")
    
    # Save raw data
    print("\nStep 5: Saving raw data...")
    results.to_csv(os.path.join(output_dir, 'monte_carlo_raw_data.csv'), index=False)
    
    # Create summary report
    summary_file = os.path.join(output_dir, 'simulation_summary.txt')
    with open(summary_file, 'w') as f:
        f.write("CONSUMER HARM MONTE CARLO SIMULATION SUMMARY\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"Simulation Date: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}\n")
        f.write(f"Platform: Railway Cloud\n")
        f.write(f"Number of simulations: {N_SIMULATIONS:,}\n")
        f.write(f"Annual transactions: {ANNUAL_TRANSACTIONS:,.0f}\n\n")
        
        for scenario_name, scenario_data in scenario_results.items():
            f.write(f"\n{scenario_name.upper()} SCENARIO\n")
            f.write("-" * 40 + "\n")
            stats = scenario_data['stats']
            f.write(f"Mean Harm: ${stats['Mean Harm']:,.2f}\n")
            f.write(f"Median Harm: ${stats['Median Harm']:,.2f}\n")
            f.write(f"95th Percentile: ${stats['95th Percentile']:,.2f}\n")
            f.write(f"99th Percentile: ${stats['99th Percentile']:,.2f}\n")
            f.write(f"Annual Industry Impact: ${stats['Annual Industry Impact (Mean)']:,.0f}\n")
            
            if scenario_name != 'Status Quo':
                reduction = (1 - stats['Mean Harm'] / scenario_results['Status Quo']['stats']['Mean Harm']) * 100
                f.write(f"Reduction from Status Quo: {reduction:.1f}%\n")
    
    print("✓ Summary report created")
    
    # Create results metadata
    metadata = {
        "simulation_date": datetime.now().isoformat(),
        "platform": "Railway",
        "simulations": N_SIMULATIONS,
        "files_generated": [
            "Consumer_Harm_Analysis.xlsx",
            "monte_carlo_raw_data.csv",
            "consumer_harm_analysis.png",
            "simulation_summary.txt"
        ],
        "key_findings": {
            "mean_harm": stats['Mean Harm'],
            "median_harm": stats['Median Harm'],
            "95th_percentile": stats['95th Percentile'],
            "annual_impact": stats['Annual Industry Impact (Mean)']
        }
    }
    
    with open(os.path.join(output_dir, 'metadata.json'), 'w') as f:
        json.dump(metadata, f, indent=2)
    
    # Create zip file for easy download
    print("\nStep 6: Creating downloadable archive...")
    zip_file = 'simulation_results.zip'
    with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, os.path.dirname(output_dir))
                zipf.write(file_path, arcname)
    
    print("✓ Archive created: simulation_results.zip")
    
    print("\n" + "="*60)
    print("SIMULATION COMPLETE!")
    print("="*60)
    print("\nGenerated files:")
    print("  ✓ Consumer_Harm_Analysis.xlsx - Comprehensive Excel report")
    print("  ✓ monte_carlo_raw_data.csv - Raw simulation data")
    print("  ✓ consumer_harm_analysis.png - Statistical visualizations")
    print("  ✓ simulation_summary.txt - Key findings summary")
    print("  ✓ simulation_results.zip - All files in one archive")
    print("\nDownload simulation_results.zip to get all your files!")
    
    return True

if __name__ == "__main__":
    try:
        success = run_complete_simulation()
        if success:
            print("\n✅ Success! The simulation completed successfully.")
            sys.exit(0)
        else:
            print("\n❌ Error: The simulation did not complete successfully.")
            sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)