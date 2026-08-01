
import pandas as pd

def format_value(x):
    return f"{x:g}" if isinstance(x,(int,float)) else x

def highlight(data,updated):
    styles=pd.DataFrame("",index=data.index,columns=data.columns)
    for item in updated:
        r=item["row"]-2
        c=item["column"]-1
        if 0<=r<len(styles.index) and 0<=c<len(styles.columns):
            styles.iloc[r,c]="background-color:#4b6043"
    return styles
