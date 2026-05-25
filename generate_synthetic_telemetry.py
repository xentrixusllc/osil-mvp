"""OSIL Enterprise Synthetic Telemetry Generator"""
import os
import random
from datetime import datetime, timedelta
import pandas as pd
import numpy as np

# Set deterministic seed so your demo looks identical every time
np.random.seed(42)
random.seed(42)

# Ensure data directory exists
os.makedirs("data", exist_ok=True)

# Time bounds (Trailing Twelve Months from March 2026)
END_DATE = datetime(2026, 3, 17)
START_DATE = END_DATE - timedelta(days=365)

def random_date(start, end):
    """Generate a random datetime between two dates"""
    delta = end - start
    int_delta = (delta.days * 24 * 60 * 60) + delta.seconds
    random_second = random.randrange(int_delta)
    return start + timedelta(seconds=random_second)

def generate_enterprise_data():
    """Generates mathematically correlated enterprise ITSM data"""
    
    # 1. Generate Problems (100 Records)
    # Weighted to create a perfect 80/20 Pareto distribution
    problem_themes = [
        ("Database Connection Pool Exhaustion", 0.40, "DBA Core"),
        ("Legacy TLS Handshake Failure", 0.25, "Network Sec"),
        ("Third Party API Timeout", 0.15, "Integration Ops"),
        ("Stale Redis Cache", 0.10, "Platform Eng"),
        ("Memory Leak", 0.05, "Platform Eng"),
        ("Hardware Node Failure", 0.05, "Data Center Ops")
    ]
    
    problems = []
    for i in range(1, 101):
        theme_data = random.choices(problem_themes, weights=[t[1] for t in problem_themes])[0]
        opened = random_date(START_DATE, END_DATE - timedelta(days=30))
        
        # 80 percent of problems are closed, 90 percent have RCA completed
        state = "Closed" if random.random() < 0.80 else "Open"
        rca_flag = 1 if random.random() < 0.90 else 0
        known_error = 1 if random.random() < 0.30 else 0
        
        # If open, it might not have an RCA text yet
        rca_text = theme_data[0] if rca_flag == 1 else ""
        
        # Map themes to specific services
        if "Database" in theme_data[0]:
            service = "Customer Portal"
        elif "TLS" in theme_data[0]:
            service = "Mobile App"
        elif "API" in theme_data[0]:
            service = "Payment Gateway"
        else:
            service = random.choice(["Inventory System", "Reporting Engine"])
            
        problems.append({
            "Service": service,
            "Problem_ID": f"PRB{1000 + i}",
            "State": state,
            "Opened_Date": opened.strftime("%Y-%m-%d %H:%M:%S"),
            "RCA_Completed_Flag": rca_flag,
            "Known_Error_Flag": known_error,
            "Root_Cause_Text": rca_text,
            "Assignment_Group": theme_data[2]
        })
        
    df_prb = pd.DataFrame(problems)
    
    # 2. Generate Changes (150 Records)
    services = ["Customer Portal", "Mobile App", "Payment Gateway", "Inventory System", "Reporting Engine"]
    changes = []
    
    for i in range(1, 151):
        svc = random.choice(services)
        start = random_date(START_DATE, END_DATE - timedelta(days=2))
        end = start + timedelta(hours=random.randint(1, 4))
        
        # Payment Gateway has high failure rate to trigger Change Collision Risk
        if svc == "Payment Gateway" and random.random() < 0.40:
            chg_type = "Emergency"
            failed = 1
            rollback = 1
        else:
            chg_type = random.choices(["Standard", "Normal", "Emergency"], weights=[0.5, 0.4, 0.1])[0]
            failed = 1 if random.random() < 0.10 else 0
            rollback = 1 if failed == 1 and random.random() < 0.50 else 0
            
        changes.append({
            "Service": svc,
            "Change_ID": f"CHG{2000 + i}",
            "Change_Type": chg_type,
            "Actual_Start": start.strftime("%Y-%m-%d %H:%M:%S"),
            "Actual_End": end.strftime("%Y-%m-%d %H:%M:%S"),
            "Change_Status": "Closed",
            "Failed_Flag": failed,
            "Rollback_Flag": rollback
        })
        
    df_chg = pd.DataFrame(changes)
    
    # 3. Generate Incidents (500 Records)
    incidents = []
    for i in range(1, 501):
        # Force the Watermelon Effect: 50 percent of all tickets go to Customer Portal as P3/P4
        svc_roll = random.random()
        if svc_roll < 0.50:
            svc = "Customer Portal"
            tier = "Tier 1"
            priority = random.choices(["P3", "P4", "P5"], weights=[0.3, 0.6, 0.1])[0]
            assign_grp = random.choice(["Service Desk", "Web Team", "DBA Core"])
            # Force Execution Churn
            reassign = random.randint(2, 6)
            reopen = 1 if random.random() < 0.25 else 0
            chg_rel = 0
            prb_id = random.choice(df_prb[df_prb["Service"] == svc]["Problem_ID"].tolist() + ["", "", ""])
            mttr_hours = random.randint(24, 120)
            
        elif svc_roll < 0.65:
            svc = "Payment Gateway"
            tier = "Tier 1"
            # Active Disruption
            priority = random.choices(["P1", "P2", "P3"], weights=[0.4, 0.4, 0.2])[0]
            assign_grp = "Integration Ops"
            reassign = random.randint(0, 2)
            reopen = 0
            chg_rel = 1 if random.random() < 0.60 else 0
            prb_id = random.choice(df_prb[df_prb["Service"] == svc]["Problem_ID"].tolist() + [""]) if len(df_prb[df_prb["Service"] == svc]) > 0 else ""
            mttr_hours = random.randint(1, 8)
            
        elif svc_roll < 0.85:
            svc = "Mobile App"
            tier = "Tier 2"
            priority = random.choices(["P2", "P3", "P4"], weights=[0.2, 0.5, 0.3])[0]
            assign_grp = "Network Sec"
            reassign = random.randint(1, 3)
            reopen = 1 if random.random() < 0.15 else 0
            chg_rel = 0
            prb_id = random.choice(df_prb[df_prb["Service"] == svc]["Problem_ID"].tolist() + ["", ""]) if len(df_prb[df_prb["Service"] == svc]) > 0 else ""
            mttr_hours = random.randint(4, 48)
            
        else:
            svc = random.choice(["Inventory System", "Reporting Engine"])
            tier = "Tier 3"
            priority = "P4"
            assign_grp = "Platform Eng"
            reassign = random.randint(0, 1)
            reopen = 0
            chg_rel = 0
            prb_id = random.choice(df_prb[df_prb["Service"] == svc]["Problem_ID"].tolist() + ["", "", ""]) if len(df_prb[df_prb["Service"] == svc]) > 0 else ""
            mttr_hours = random.randint(2, 24)

        opened = random_date(START_DATE, END_DATE - timedelta(days=5))
        resolved = opened + timedelta(hours=mttr_hours)
        
        incidents.append({
            "Service": svc,
            "Service_Tier": tier,
            "Opened_Date": opened.strftime("%Y-%m-%d %H:%M:%S"),
            "Resolved_Date": resolved.strftime("%Y-%m-%d %H:%M:%S"),
            "State": "Closed",
            "Priority": priority,
            "Assignment_Group": assign_grp,
            "Reassignment_Count": reassign,
            "Reopened_Flag": reopen,
            "Change_Related_Flag": chg_rel,
            "Problem_ID": prb_id
        })
        
    df_inc = pd.DataFrame(incidents)
    
    # 4. Save to CSV in the data directory
    df_prb.to_csv("data/demo_problems.csv", index=False)
    df_chg.to_csv("data/demo_changes.csv", index=False)
    df_inc.to_csv("data/demo_incidents.csv", index=False)
    
    print("SUCCESS: Enterprise synthetic telemetry generated. 500 Incidents, 150 Changes, 100 Problems.")
    print("Files saved to data/ directory.")

if __name__ == "__main__":
    generate_enterprise_data()
