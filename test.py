import pandas as pd

def test_file():
    file_path = "bank_statements/תזרים-יומי-תצוגה-מפורטת-60-הימים-הקרובים-רומי-רפאל-נדלן-בעמ.xlsx"
    
    print("קורא את הקובץ...")
    df = pd.read_excel(file_path, header=3)
    
    print("\nכותרות בקובץ:")
    print(df.columns.tolist())
    
    print("\n3 שורות ראשונות:")
    print(df[['תאריך', 'ח-ן', 'תיאור', 'חובה', 'זכות', 'יתרה (כולל נגררות)']].head(3))
    
    print("\nסוגי הנתונים:")
    print(df.dtypes)

if __name__ == "__main__":
    test_file()