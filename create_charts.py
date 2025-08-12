import json
import matplotlib.pyplot as plt
import numpy as np
import os
from collections import Counter

def create_charts():
    
    try:
        with open('pdf_extraction_result.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        if not os.path.exists('charts'):
            os.makedirs('charts')
            
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
        
        division_counter = Counter(row['division'] for row in all_rows)
        top_divisions = division_counter.most_common(10)
        
        group_counter = Counter(row['group'] for row in all_rows)
        top_groups = group_counter.most_common(10)
        
        department_counter = Counter(row['department'] for row in all_rows)
        top_departments = department_counter.most_common(10)
        
        div_group_pairs = Counter((row['division'], row['group']) for row in all_rows)
        top_pairs = div_group_pairs.most_common(10)
        
        persons_per_division = {}
        for row in all_rows:
            if row['person'] not in persons_per_division.get(row['division'], set()):
                div_set = persons_per_division.get(row['division'], set())
                div_set.add(row['person'])
                persons_per_division[row['division']] = div_set
                
        top_divisions_by_persons = sorted(
            [(div, len(persons)) for div, persons in persons_per_division.items()],
            key=lambda x: x[1], 
            reverse=True
        )[:10]
        
        plt.figure(figsize=(12, 6))
        divisions = [d[0] for d in top_divisions]
        counts = [d[1] for d in top_divisions]
        
        plt.bar(divisions, counts, color='skyblue')
        plt.xlabel('Division')
        plt.ylabel('Number of Occurrences')
        plt.title('Top 10 Divisions by Occurrence')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.savefig('charts/division_distribution.png')
        plt.close()
        
        plt.figure(figsize=(12, 6))
        groups = [g[0] for g in top_groups]
        counts = [g[1] for g in top_groups]
        
        plt.bar(groups, counts, color='lightgreen')
        plt.xlabel('Group')
        plt.ylabel('Number of Occurrences')
        plt.title('Top 10 Groups by Occurrence')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.savefig('charts/group_distribution.png')
        plt.close()
        
        plt.figure(figsize=(12, 6))
        departments = [d[0] for d in top_departments]
        counts = [d[1] for d in top_departments]
        
        plt.bar(departments, counts, color='salmon')
        plt.xlabel('Department')
        plt.ylabel('Number of Occurrences')
        plt.title('Top 10 Departments by Occurrence')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.savefig('charts/department_distribution.png')
        plt.close()
        
        plt.figure(figsize=(10, 10))
        divisions = [d[0] for d in top_divisions_by_persons[:5]]
        counts = [d[1] for d in top_divisions_by_persons[:5]]
        others = sum(d[1] for d in top_divisions_by_persons[5:])
        
        divisions.append('Others')
        counts.append(others)
        
        plt.pie(counts, labels=divisions, autopct='%1.1f%%', startangle=90, shadow=True)
        plt.axis('equal')
        plt.title('Resource Distribution by Division')
        plt.tight_layout()
        plt.savefig('charts/resource_distribution_pie.png')
        plt.close()
        
        plt.figure(figsize=(14, 7))
        pairs = [f"{p[0][0]}-{p[0][1]}" for p in top_pairs[:7]]
        counts = [p[1] for p in top_pairs[:7]]
        
        plt.bar(pairs, counts, color='mediumpurple')
        plt.xlabel('Division-Group Pair')
        plt.ylabel('Number of Occurrences')
        plt.title('Top Division-Group Relationships')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.savefig('charts/division_group_relationships.png')
        plt.close()
        
        print("Charts created successfully in the 'charts' directory!")
        print("Use these charts in your PowerPoint presentation.")
        
    except Exception as e:
        print(f"Error creating charts: {str(e)}")

if __name__ == "__main__":
    create_charts() 