#!/bin/bash

# Extract Access Database Schema and Data
# Usage: ./extract-access-data.sh "path/to/database.accdb"

if [ $# -eq 0 ]; then
    echo "Usage: $0 <path-to-access-database>"
    exit 1
fi

DATABASE="$1"
OUTPUT_DIR="data"
VBA_DIR="vba-export"

# Check if database exists
if [ ! -f "$DATABASE" ]; then
    echo "Error: Database file not found: $DATABASE"
    exit 1
fi

# Check if mdbtools is installed
if ! command -v mdb-tables &> /dev/null; then
    echo "Error: mdbtools is not installed"
    echo "Install with: sudo apt-get install mdbtools (Ubuntu/Debian)"
    echo "             or: brew install mdbtools (Mac)"
    exit 1
fi

# Create output directories
mkdir -p "$OUTPUT_DIR/tables"
mkdir -p "$OUTPUT_DIR/queries"
mkdir -p "$VBA_DIR"

echo "========================================="
echo "Extracting Access Database"
echo "========================================="
echo "Database: $DATABASE"
echo ""

# Extract schema
echo "Extracting schema..."
mdb-schema "$DATABASE" mysql > "$OUTPUT_DIR/schema.sql"
echo "✓ Schema exported to: $OUTPUT_DIR/schema.sql"

# Get list of tables
echo ""
echo "Extracting tables..."
tables=$(mdb-tables -1 "$DATABASE")

# Export each table to CSV
for table in $tables; do
    echo "  - Exporting: $table"
    mdb-export "$DATABASE" "$table" > "$OUTPUT_DIR/tables/${table}.csv"
done
echo "✓ Tables exported to: $OUTPUT_DIR/tables/"

# Count records in each table
echo ""
echo "Table Statistics:"
echo "----------------------------------------"
for table in $tables; do
    count=$(wc -l < "$OUTPUT_DIR/tables/${table}.csv")
    count=$((count - 1))  # Subtract header row
    printf "%-30s %10d rows\n" "$table" "$count"
done

# Create documentation file
echo ""
echo "Generating documentation..."
cat > "$OUTPUT_DIR/database-info.md" << EOF
# Database Information

**File:** $DATABASE
**Extracted:** $(date)

## Tables

$(for table in $tables; do
    count=$(wc -l < "$OUTPUT_DIR/tables/${table}.csv")
    count=$((count - 1))
    echo "- **$table** ($count records)"
done)

## Schema

See \`schema.sql\` for complete database schema.

## Data Files

All tables exported to CSV format in \`tables/\` directory.

EOF

echo "✓ Documentation created: $OUTPUT_DIR/database-info.md"

echo ""
echo "========================================="
echo "Extraction Complete!"
echo "========================================="
echo ""
echo "Next steps:"
echo "1. Review extracted data in: $OUTPUT_DIR/"
echo "2. Update docs/02-DATABASE-SCHEMA.md with table information"
echo "3. For VBA code, use Microsoft Access or see scripts/README.md"
echo ""
