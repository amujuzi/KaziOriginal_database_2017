#!/usr/bin/env python3
"""
Extract schema and data from MS Access database
Requires: pip install pyodbc pandas

Usage: python extract-access-data.py "path/to/database.accdb"
"""

import sys
import os
from pathlib import Path
import pyodbc
import pandas as pd
from datetime import datetime

def get_connection_string(database_path):
    """Get ODBC connection string for Access database"""
    abs_path = os.path.abspath(database_path)

    # Try different drivers
    drivers = [
        'Microsoft Access Driver (*.mdb, *.accdb)',
        'Microsoft Access Driver (*.mdb)',
        'MICROSOFT ACCESS DRIVER (*.mdb, *.accdb)',
    ]

    for driver in drivers:
        try:
            conn_str = f'DRIVER={{{driver}}};DBQ={abs_path};'
            conn = pyodbc.connect(conn_str)
            conn.close()
            return conn_str
        except:
            continue

    raise Exception("No compatible Access driver found. Install Microsoft Access Database Engine.")

def extract_database(database_path):
    """Extract all tables and metadata from Access database"""

    print("=" * 60)
    print("Extracting Access Database")
    print("=" * 60)
    print(f"Database: {database_path}\n")

    # Create output directories
    output_dir = Path("data")
    tables_dir = output_dir / "tables"
    tables_dir.mkdir(parents=True, exist_ok=True)

    vba_dir = Path("vba-export")
    vba_dir.mkdir(exist_ok=True)

    # Connect to database
    try:
        conn_str = get_connection_string(database_path)
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    except Exception as e:
        print(f"Error connecting to database: {e}")
        print("\nMake sure you have Microsoft Access Database Engine installed:")
        print("https://www.microsoft.com/en-us/download/details.aspx?id=54920")
        return

    # Get list of tables
    tables = []
    for table_info in cursor.tables(tableType='TABLE'):
        table_name = table_info.table_name
        # Skip system tables
        if not table_name.startswith('MSys') and not table_name.startswith('~'):
            tables.append(table_name)

    print(f"Found {len(tables)} tables\n")
    print("Extracting tables...")

    table_stats = []

    # Export each table
    for table_name in tables:
        try:
            # Read table data
            query = f"SELECT * FROM [{table_name}]"
            df = pd.read_sql(query, conn)

            # Export to CSV
            csv_path = tables_dir / f"{table_name}.csv"
            df.to_csv(csv_path, index=False)

            row_count = len(df)
            table_stats.append((table_name, row_count, len(df.columns)))

            print(f"  ✓ {table_name}: {row_count} rows, {len(df.columns)} columns")

        except Exception as e:
            print(f"  ✗ Error exporting {table_name}: {e}")

    # Get table schemas
    print("\nExtracting table schemas...")
    schema_info = []

    for table_name in tables:
        try:
            cursor.execute(f"SELECT * FROM [{table_name}] WHERE 1=0")
            columns = []
            for col in cursor.description:
                columns.append({
                    'name': col[0],
                    'type': col[1].__name__ if col[1] else 'unknown',
                    'size': col[3] if col[3] else 0
                })
            schema_info.append((table_name, columns))
        except Exception as e:
            print(f"  Error getting schema for {table_name}: {e}")

    # Create documentation
    print("\nGenerating documentation...")

    doc_content = f"""# Database Information

**File:** {database_path}
**Extracted:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Table Summary

| Table Name | Rows | Columns |
|------------|------|---------|
"""

    for table_name, row_count, col_count in sorted(table_stats):
        doc_content += f"| {table_name} | {row_count:,} | {col_count} |\n"

    doc_content += "\n## Table Schemas\n\n"

    for table_name, columns in schema_info:
        doc_content += f"### {table_name}\n\n"
        doc_content += "| Column Name | Data Type | Size |\n"
        doc_content += "|-------------|-----------|------|\n"
        for col in columns:
            doc_content += f"| {col['name']} | {col['type']} | {col['size']} |\n"
        doc_content += "\n"

    # Save documentation
    doc_path = output_dir / "database-info.md"
    with open(doc_path, 'w') as f:
        f.write(doc_content)

    print(f"✓ Documentation saved to: {doc_path}")

    # Create Laravel migrations template
    print("\nGenerating Laravel migration templates...")
    migrations_dir = output_dir / "laravel-migrations"
    migrations_dir.mkdir(exist_ok=True)

    for table_name, columns in schema_info:
        migration_content = generate_laravel_migration(table_name, columns)
        timestamp = datetime.now().strftime('%Y_%m_%d_%H%M%S')
        migration_file = migrations_dir / f"{timestamp}_create_{table_name.lower()}_table.php"
        with open(migration_file, 'w') as f:
            f.write(migration_content)

    print(f"✓ Migration templates saved to: {migrations_dir}")

    # Close connection
    conn.close()

    print("\n" + "=" * 60)
    print("Extraction Complete!")
    print("=" * 60)
    print("\nNext steps:")
    print("1. Review extracted data in: data/")
    print("2. Check generated migrations in: data/laravel-migrations/")
    print("3. Update docs/02-DATABASE-SCHEMA.md with table information")
    print("4. For VBA code extraction, use Microsoft Access")
    print()

def generate_laravel_migration(table_name, columns):
    """Generate Laravel migration template"""

    # Convert Access types to Laravel migration types
    type_map = {
        'int': 'integer',
        'str': 'string',
        'float': 'decimal',
        'bool': 'boolean',
        'datetime': 'dateTime',
        'date': 'date',
        'bytes': 'binary'
    }

    snake_table = table_name.lower().replace('tbl', '').replace('_', '')

    migration = f"""<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{{
    public function up(): void
    {{
        Schema::create('{snake_table}', function (Blueprint $table) {{
            $table->id();
"""

    for col in columns:
        if col['name'].lower() in ['id', 'created_at', 'updated_at']:
            continue

        col_type = type_map.get(col['type'], 'string')
        col_name = col['name'].lower()

        if col_type == 'string' and col['size'] > 0:
            migration += f"            $table->string('{col_name}', {col['size']});\n"
        else:
            migration += f"            $table->{col_type}('{col_name}');\n"

    migration += """            $table->timestamps();
        });
    }

    public function down(): void
    {
        Schema::dropIfExists('""" + snake_table + """');
    }
};
"""

    return migration

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python extract-access-data.py <path-to-database>")
        sys.exit(1)

    database_path = sys.argv[1]

    if not os.path.exists(database_path):
        print(f"Error: Database file not found: {database_path}")
        sys.exit(1)

    try:
        extract_database(database_path)
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
