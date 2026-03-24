"""OSIL Persistent Tenant History Ledger"""
import sqlite3
import os
from datetime import datetime
from typing import Dict

import pandas as pd

DB_PATH = "osil_tenant_history.db"

def init_db() -> None:
    """Initialize the SQLite database and create the history table if it does not exist."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tenant_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tenant_name TEXT NOT NULL,
            run_date TEXT NOT NULL,
            bvsi_score REAL,
            resilience_score REAL,
            governance_score REAL,
            debt_score REAL,
            momentum_score REAL,
            execution_timestamp TEXT NOT NULL
        )
    ''')
    
    conn.commit()
    conn.close()

def save_tenant_run(tenant_name: str, as_of_date: str, bvsi: float, domain_scores: Dict[str, float]) -> None:
    """Insert a new monthly execution record for a specific tenant."""
    init_db() 
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    resilience = float(domain_scores.get("Service Resilience", 0.0))
    governance = float(domain_scores.get("Change Governance", 0.0))
    debt = float(domain_scores.get("Structural Risk Debt", domain_scores.get("Structural Risk Debt™", 0.0)))
    momentum = float(domain_scores.get("Reliability Momentum", 0.0))
    
    now = datetime.now().isoformat()
    
    cursor.execute('''
        INSERT INTO tenant_history 
        (tenant_name, run_date, bvsi_score, resilience_score, governance_score, debt_score, momentum_score, execution_timestamp)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (tenant_name, as_of_date, bvsi, resilience, governance, debt, momentum, now))
    
    conn.commit()
    conn.close()

def get_tenant_history(tenant_name: str) -> pd.DataFrame:
    """Retrieve the chronological execution history for a specific tenant to power trendline visuals."""
    if not os.path.exists(DB_PATH):
        return pd.DataFrame()
        
    conn = sqlite3.connect(DB_PATH)
    
    try:
        query = """
            SELECT run_date, bvsi_score, resilience_score, governance_score, debt_score, momentum_score 
            FROM tenant_history 
            WHERE tenant_name = ? 
            ORDER BY run_date ASC
        """
        df = pd.read_sql_query(query, conn, params=(tenant_name,))
    except sqlite3.OperationalError:
        df = pd.DataFrame()
        
    conn.close()
    
    if not df.empty:
        df["run_date"] = pd.to_datetime(df["run_date"]).dt.strftime('%Y-%m-%d')
        df = df.drop_duplicates(subset=["run_date"], keep="last")
        
    return df

def get_all_tenants() -> list:
    """Retrieve a list of unique tenant names to populate the workspace dropdown."""
    if not os.path.exists(DB_PATH):
        return []
        
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        query = "SELECT DISTINCT tenant_name FROM tenant_history ORDER BY tenant_name ASC"
        cursor.execute(query)
        rows = cursor.fetchall()
    except sqlite3.OperationalError:
        rows = []
    
    conn.close()
    
    return [row[0] for row in rows]
