#!/usr/bin/env python3
"""
Consumer Harm Monte Carlo Simulation
Analyzes potential consumer harm across various scenarios including hidden fees,
service failures, and damages with claim denials.
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings('ignore')

# Set random seed for reproducibility
np.random.seed(42)

# Configuration
N_SIMULATIONS = 10000  # Number of simulated customers
ANNUAL_TRANSACTIONS = 1.73e6  # 1.73 million transactions per year

# Distribution parameters (using triangular distributions)
PARAMS = {
    'base_service_cost': {'min': 2500, 'mode': 3200, 'max': 4000},
    'hidden_fees': {'min': 0, 'mode': 375, 'max': 1100},
    'service_failure_prob': {'min': 0.15, 'mode': 0.30, 'max': 0.45},
    'claim_denial_prob': {'min': 0.60, 'mode': 0.85, 'max': 0.95},
    'damage_occurrence_rate': {'min': 0.05, 'mode': 0.12, 'max': 0.25},
    'average_damage_value': {'min': 500, 'mode': 2500, 'max': 10000}
}

# Penalty cost for service failures (assumption)
SERVICE_FAILURE_PENALTY = 1000

def triangular_sample(min_val, mode_val, max_val, size):
    """Generate samples from triangular distribution"""
    return np.random.triangular(min_val, mode_val, max_val, size)

def run_monte_carlo_simulation(params=PARAMS, n_sims=N_SIMULATIONS):
    """
    Run Monte Carlo simulation for consumer harm
    """
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
        'Customers with Harm > $5000': (harm > 5000).sum()
    }
    
    # Industry-wide annual impact
    stats_dict['Annual Industry Impact (Mean)'] = stats_dict['Mean Harm'] * ANNUAL_TRANSACTIONS
    stats_dict['Annual Industry Impact (95th %ile)'] = stats_dict['95th Percentile'] * ANNUAL_TRANSACTIONS
    
    return stats_dict

def create_visualizations(results):
    """Create comprehensive visualizations"""
    
    # Set style
    plt.style.use('seaborn-v0_8-darkgrid')
    
    # Create figure with subplots
    fig = plt.figure(figsize=(20, 12))
    
    # 1. Histogram of total harm
    ax1 = plt.subplot(2, 3, 1)
    plt.hist(results['total_harm'], bins=50, color='steelblue', alpha=0.7, edgecolor='black')
    plt.axvline(results['total_harm'].mean(), color='red', linestyle='--', linewidth=2, label=f'Mean: ${results["total_harm"].mean():.0f}')
    plt.axvline(results['total_harm'].median(), color='green', linestyle='--', linewidth=2, label=f'Median: ${results["total_harm"].median():.0f}')
    plt.xlabel('Total Consumer Harm ($)')
    plt.ylabel('Frequency')
    plt.title('Distribution of Consumer Harm')
    plt.legend()
    
    # 2. Box plot of harm components
    ax2 = plt.subplot(2, 3, 2)
    harm_components = pd.DataFrame({
        'Hidden Fees': results['hidden_fees'],
        'Service Failures': results['service_failure_harm'],
        'Damages (Denied Claims)': results['damage_harm']
    })
    harm_components.boxplot(ax=ax2)
    plt.ylabel('Harm Amount ($)')
    plt.title('Harm Components Distribution')
    plt.xticks(rotation=45)
    
    # 3. Cumulative distribution
    ax3 = plt.subplot(2, 3, 3)
    sorted_harm = np.sort(results['total_harm'])
    cumulative = np.arange(1, len(sorted_harm) + 1) / len(sorted_harm) * 100
    plt.plot(sorted_harm, cumulative, linewidth=2, color='darkblue')
    plt.xlabel('Total Consumer Harm ($)')
    plt.ylabel('Cumulative Percentage (%)')
    plt.title('Cumulative Distribution of Consumer Harm')
    plt.grid(True, alpha=0.3)
    
    # 4. Harm by component pie chart
    ax4 = plt.subplot(2, 3, 4)
    component_means = [
        results['hidden_fees'].mean(),
        results['service_failure_harm'].mean(),
        results['damage_harm'].mean()
    ]
    labels = ['Hidden Fees', 'Service Failures', 'Damages (Denied Claims)']
    colors = ['#ff9999', '#66b3ff', '#99ff99']
    plt.pie(component_means, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
    plt.title('Average Harm Breakdown by Component')
    
    # 5. Scatter plot: Service Cost vs Total Harm
    ax5 = plt.subplot(2, 3, 5)
    scatter = plt.scatter(results['service_cost'], results['total_harm'], 
                         c=results['hidden_fees'], cmap='viridis', alpha=0.5, s=10)
    plt.xlabel('Base Service Cost ($)')
    plt.ylabel('Total Consumer Harm ($)')
    plt.title('Service Cost vs Total Harm (colored by Hidden Fees)')
    plt.colorbar(scatter, label='Hidden Fees ($)')
    
    # 6. Percentile chart
    ax6 = plt.subplot(2, 3, 6)
    percentiles = [10, 25, 50, 75, 90, 95, 99]
    percentile_values = [results['total_harm'].quantile(p/100) for p in percentiles]
    plt.bar([str(p) + 'th' for p in percentiles], percentile_values, color='coral')
    plt.xlabel('Percentile')
    plt.ylabel('Harm Amount ($)')
    plt.title('Consumer Harm by Percentile')
    for i, v in enumerate(percentile_values):
        plt.text(i, v + 50, f'${v:.0f}', ha='center', va='bottom')
    
    plt.tight_layout()
    plt.savefig('consumer_harm_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()

def create_interactive_visualizations(results):
    """Create interactive Plotly visualizations"""
    
    # Create subplots
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=('Harm Distribution', 'Harm Components', 
                       'Cumulative Distribution', 'Scenario Comparison'),
        specs=[[{'type': 'histogram'}, {'type': 'box'}],
               [{'type': 'scatter'}, {'type': 'bar'}]]
    )
    
    # 1. Interactive histogram
    fig.add_trace(
        go.Histogram(x=results['total_harm'], nbinsx=50, name='Total Harm',
                    marker_color='steelblue'),
        row=1, col=1
    )
    
    # 2. Box plots for components
    for component, color in zip(['hidden_fees', 'service_failure_harm', 'damage_harm'],
                               ['red', 'blue', 'green']):
        fig.add_trace(
            go.Box(y=results[component], name=component.replace('_', ' ').title(),
                   marker_color=color),
            row=1, col=2
        )
    
    # 3. Cumulative distribution
    sorted_harm = np.sort(results['total_harm'])
    cumulative = np.arange(1, len(sorted_harm) + 1) / len(sorted_harm) * 100
    fig.add_trace(
        go.Scatter(x=sorted_harm, y=cumulative, mode='lines', name='Cumulative %',
                  line=dict(color='darkblue', width=2)),
        row=2, col=1
    )
    
    # 4. Scenario comparison (example with different assumptions)
    scenarios = ['Status Quo', 'Moderate Reform', 'Strong Reform']
    mean_harms = [results['total_harm'].mean(), 
                  results['total_harm'].mean() * 0.6,  # 40% reduction
                  results['total_harm'].mean() * 0.3]   # 70% reduction
    
    fig.add_trace(
        go.Bar(x=scenarios, y=mean_harms, name='Mean Harm by Scenario',
               marker_color=['red', 'orange', 'green']),
        row=2, col=2
    )
    
    # Update layout
    fig.update_layout(height=800, showlegend=True,
                     title_text="Consumer Harm Monte Carlo Analysis Dashboard")
    fig.update_xaxes(title_text="Total Harm ($)", row=1, col=1)
    fig.update_xaxes(title_text="Total Harm ($)", row=2, col=1)
    fig.update_yaxes(title_text="Frequency", row=1, col=1)
    fig.update_yaxes(title_text="Harm ($)", row=1, col=2)
    fig.update_yaxes(title_text="Cumulative %", row=2, col=1)
    fig.update_yaxes(title_text="Mean Harm ($)", row=2, col=2)
    
    # Save as HTML
    fig.write_html("consumer_harm_interactive.html")
    fig.show()

def run_scenario_analysis():
    """Run analysis for different scenarios"""
    
    scenarios = {
        'Status Quo': PARAMS,
        'Moderate Reform': {
            'base_service_cost': PARAMS['base_service_cost'],
            'hidden_fees': {'min': 0, 'mode': 150, 'max': 500},  # Reduced hidden fees
            'service_failure_prob': {'min': 0.10, 'mode': 0.20, 'max': 0.30},  # Better service
            'claim_denial_prob': {'min': 0.40, 'mode': 0.60, 'max': 0.80},  # Fairer claims
            'damage_occurrence_rate': PARAMS['damage_occurrence_rate'],
            'average_damage_value': PARAMS['average_damage_value']
        },
        'Strong Reform': {
            'base_service_cost': PARAMS['base_service_cost'],
            'hidden_fees': {'min': 0, 'mode': 50, 'max': 200},  # Minimal hidden fees
            'service_failure_prob': {'min': 0.05, 'mode': 0.10, 'max': 0.15},  # Excellent service
            'claim_denial_prob': {'min': 0.20, 'mode': 0.35, 'max': 0.50},  # Fair claims
            'damage_occurrence_rate': {'min': 0.03, 'mode': 0.08, 'max': 0.15},  # Better handling
            'average_damage_value': PARAMS['average_damage_value']
        }
    }
    
    scenario_results = {}
    
    for scenario_name, scenario_params in scenarios.items():
        print(f"\nRunning scenario: {scenario_name}")
        results = run_monte_carlo_simulation(scenario_params)
        stats = calculate_statistics(results)
        scenario_results[scenario_name] = {
            'results': results,
            'stats': stats
        }
        
        print(f"Mean Harm: ${stats['Mean Harm']:,.2f}")
        print(f"95th Percentile: ${stats['95th Percentile']:,.2f}")
        print(f"Annual Industry Impact: ${stats['Annual Industry Impact (Mean)']:,.0f}")
    
    return scenario_results

def main():
    """Main execution function"""
    print("Consumer Harm Monte Carlo Simulation")
    print("=" * 50)
    
    # Run base simulation
    print("\nRunning base simulation with {} iterations...".format(N_SIMULATIONS))
    results = run_monte_carlo_simulation()
    
    # Calculate statistics
    stats = calculate_statistics(results)
    
    # Print key statistics
    print("\nKey Statistics:")
    print("-" * 30)
    for key, value in stats.items():
        if 'Annual' in key:
            print(f"{key}: ${value:,.0f}")
        elif isinstance(value, (int, float)) and value > 100:
            print(f"{key}: ${value:,.2f}")
        else:
            print(f"{key}: {value:,.0f}")
    
    # Create visualizations
    print("\nGenerating visualizations...")
    create_visualizations(results)
    create_interactive_visualizations(results)
    
    # Run scenario analysis
    print("\nRunning scenario analysis...")
    scenario_results = run_scenario_analysis()
    
    # Export results to CSV
    results.to_csv('monte_carlo_results.csv', index=False)
    
    # Create summary report
    with open('simulation_summary.txt', 'w') as f:
        f.write("Consumer Harm Monte Carlo Simulation Summary\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"Number of simulations: {N_SIMULATIONS:,}\n")
        f.write(f"Annual transactions: {ANNUAL_TRANSACTIONS:,.0f}\n\n")
        
        for scenario_name, scenario_data in scenario_results.items():
            f.write(f"\n{scenario_name} Scenario:\n")
            f.write("-" * 30 + "\n")
            for key, value in scenario_data['stats'].items():
                if 'Annual' in key:
                    f.write(f"{key}: ${value:,.0f}\n")
                elif isinstance(value, (int, float)) and value > 100:
                    f.write(f"{key}: ${value:,.2f}\n")
                else:
                    f.write(f"{key}: {value:,.0f}\n")
    
    print("\nSimulation complete! Files generated:")
    print("- consumer_harm_analysis.png (static visualizations)")
    print("- consumer_harm_interactive.html (interactive dashboard)")
    print("- monte_carlo_results.csv (raw simulation data)")
    print("- simulation_summary.txt (summary statistics)")

if __name__ == "__main__":
    main()
