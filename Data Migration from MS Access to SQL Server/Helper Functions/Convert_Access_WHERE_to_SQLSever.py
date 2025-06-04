import re

def convert_access_WHERE_to_sqlserver(sql: str):
    sql = sql.replace('"', "'")
    matches = re.findall(r"LIKE\s*'[^']+'", sql, flags=re.IGNORECASE)

    for match in matches:
        original = match
        pattern = re.search(r"'([^']+)'", match)  

        if pattern:
            result = pattern.group(1) 
        else:
            result = None

        converted_result = result.replace('*', '%')
        sql = sql.replace(original, f"LIKE '{converted_result}'")
    
    return sql

access_sql = ''' 
WHERE (((ROOM.Activity) Like "*toil*" Or (ROOM.Activity) Like "*restroom*" Or (ROOM.Activity) Like "restrm*" Or (ROOM.Activity) Like "*locker*" Or (ROOM.Activity) Like "*shower*" Or (ROOM.Activity) Like "*tub*" Or (ROOM.Activity) Like "*bath*"));
'''

print(convert_access_WHERE_to_sqlserver(access_sql))