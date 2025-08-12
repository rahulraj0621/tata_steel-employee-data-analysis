import json
import os
from collections import Counter

def analyze_pdf_data():
    
    try:
        with open('pdf_extraction_result.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        total_pages = data['summary']['total_pages']
        pages_with_content = data['summary']['pages_with_content']
        
        print(f"Total pages: {total_pages}")
        print(f"Pages with content: {pages_with_content}")
        
        all_rows = []
        for page in data['content']:
            content = page['content']
            lines = content.strip().split('\n')
            for line in lines:
                if line.startswith('Person_'):
                    parts = line.strip().split()
                    if len(parts) >= 4:  
                        all_rows.append({
                            'person': parts[0],
                            'division': parts[1],
                            'group': parts[2],
                            'department': parts[3]
                        })
        
        unique_persons = set(row['person'] for row in all_rows)
        unique_divisions = set(row['division'] for row in all_rows)
        unique_groups = set(row['group'] for row in all_rows)
        unique_departments = set(row['department'] for row in all_rows)
        
        print(f"\nUnique Persons: {len(unique_persons)}")
        print(f"Unique Divisions: {len(unique_divisions)}")
        print(f"Unique Groups: {len(unique_groups)}")
        print(f"Unique Departments: {len(unique_departments)}")
        
        division_counter = Counter(row['division'] for row in all_rows)
        print("\nTop Divisions:")
        for division, count in division_counter.most_common(10):
            print(f"  {division}: {count} occurrences")
            
        group_counter = Counter(row['group'] for row in all_rows)
        print("\nTop Groups:")
        for group, count in group_counter.most_common(10):
            print(f"  {group}: {count} occurrences")
            
        department_counter = Counter(row['department'] for row in all_rows)
        print("\nTop Departments:")
        for dept, count in department_counter.most_common(10):
            print(f"  {dept}: {count} occurrences")
            
        div_group_pairs = Counter((row['division'], row['group']) for row in all_rows)
        print("\nTop Division-Group Pairs:")
        for pair, count in div_group_pairs.most_common(10):
            print(f"  {pair[0]} - {pair[1]}: {count} occurrences")
            
        persons_per_division = {}
        for row in all_rows:
            if row['person'] not in persons_per_division.get(row['division'], set()):
                div_set = persons_per_division.get(row['division'], set())
                div_set.add(row['person'])
                persons_per_division[row['division']] = div_set
                
        print("\nPersons per Division:")
        for division, persons in sorted(persons_per_division.items(), key=lambda x: len(x[1]), reverse=True)[:10]:
            print(f"  {division}: {len(persons)} unique persons")
            
        with open('data_analysis_summary.txt', 'w', encoding='utf-8') as f:
            f.write("# Tata Steel Data Analysis Summary\n\n")
            f.write(f"Total records: {len(all_rows)}\n")
            f.write(f"Unique persons: {len(unique_persons)}\n")
            f.write(f"Unique divisions: {len(unique_divisions)}\n")
            f.write(f"Unique groups: {len(unique_groups)}\n")
            f.write(f"Unique departments: {len(unique_departments)}\n\n")
            
            f.write("## Top Divisions\n")
            for division, count in division_counter.most_common(5):
                f.write(f"- {division}: {count} occurrences\n")
                
            f.write("\n## Top Groups\n")
            for group, count in group_counter.most_common(5):
                f.write(f"- {group}: {count} occurrences\n")
                
            f.write("\n## Top Departments\n")
            for dept, count in department_counter.most_common(5):
                f.write(f"- {dept}: {count} occurrences\n")
                
            f.write("\n## Division-Group Distribution\n")
            for pair, count in div_group_pairs.most_common(5):
                f.write(f"- {pair[0]} with {pair[1]}: {count} occurrences\n")
                
            f.write("\n## Resource Distribution\n")
            for division, persons in sorted(persons_per_division.items(), key=lambda x: len(x[1]), reverse=True)[:5]:
                f.write(f"- {division}: {len(persons)} unique persons\n")
        
        print("\nAnalysis complete! Summary saved to 'data_analysis_summary.txt'")
        
    except Exception as e:
        print(f"Error analyzing data: {str(e)}")

if __name__ == "__main__":
    analyze_pdf_data() 