import streamlit as st
import pandas as pd

#Load data from Excel file with multiple sheets
@st.cache_data
def load_data(file_path):
#xls = pd.ExcelFile("C:/Users/davood_shahbakhti/Downloads/Projects Status/ProjectsStatus.xls")
    file_path = "ProjectsStatus.xlsx"
    xls = pd.ExcelFile("ProjectsStatus.xlsx")
    projects = pd.read_excel(xls, 'Projects')
    wbs = pd.read_excel(xls, 'WBS')
    activities = pd.read_excel(xls, 'Activities')
    resources = pd.read_excel(xls, 'Resources')
    assignments = pd.read_excel(xls, 'Assignments')
    procurement = pd.read_excel(xls, 'Procurement')
    installation = pd.read_excel(xls, 'Installation')
    return projects, wbs, activities, resources, assignments, procurement, installation

# culations and data processing
def preprocess_data(projects, wbs, activities, resources, assignments, procurement, installation):
    # Convert date columns to datetime
    projects['تاریخ شروع'] = pd.to_datetime(projects['تاریخ شروع'])
    projects['تاریخ پایان'] = pd.to_datetime(projects['تاریخ پایان'])
    activities['تاریخ شروع'] = pd.to_datetime(activities['تاریخ شروع'])
    activities['تاریخ پایان'] = pd.to_datetime(activities['تاریخ پایان'])
    procurement['تاریخ تحویل'] = pd.to_datetime(procurement['تاریخ تحویل'])
    installation['تاریخ نصب'] = pd.to_datetime(installation['تاریخ نصب'])

    # Merge assignments with resources
    assign_res = assignments.merge(resources, left_on='کد منبع', right_on='کد منبع', how='left')
    # Calculate actual vs planned hours difference
    assign_res['ساعت تفاوت'] = assign_res['ساعت واقعی'] - assign_res['ساعت برنامه‌ریزی‌شده']
    # Sum costs per activity
    activity_costs = assign_res.groupby('کد فعالیت')[['هزینه کل (میلیون ریال)']].sum().reset_index()
    activities = activities.merge(activity_costs, left_on='کد فعالیت', right_on='کد فعالیت', how='left')
    activities.rename(columns={'هزینه کل (میلیون ریال)': 'هزینه تخصیص یافته'}, inplace=True)

    return projects, activities, assign_res
def main():
    st.set_page_config(page_title="پروتوتایپ داشبورد پروژه", layout="wide")

    st.title("داشبورد پروژه - Primavera P6 و SAP PS")

    uploaded_file = st.file_uploader("لطفاً فایل اکسل داده‌های پروژه را بارگذاری کنید", type=["xlsx"])
    if uploaded_file is None:
        st.warning("برای شروع، لطفاً فایل اکسل داده‌های پروژه را بارگذاری کنید.")
        return

    projects, wbs, activities, resources, assignments, procurement, installation = load_data(uploaded_file)
    projects, activities, assign_res = preprocess_data(projects, activities, assignments, resources)

    # نمایش خلاصه پروژه‌ها
    st.header("خلاصه پروژه‌ها")
    st.dataframe(projects[['شناسه پروژه', 'نام پروژه', 'مدیر پروژه', 'تاریخ شروع', 'تاریخ پایان', 'درصد پیشرفت', 'وضعیت']])

    # نمایش فعالیت‌ها
    st.header("فعالیت‌ها")
    st.dataframe(activities[['کد فعالیت', 'نام فعالیت', 'تاریخ شروع', 'تاریخ پایان', 'مدت (روز)', 'درصد پیشرفت', 'هزینه برنامه‌ریزی‌شده (میلیون ریال)', 'هزینه واقعی (میلیون ریال)', 'هزینه تخصیص یافته', 'وضعیت']])

    # نمایش منابع
    st.header("منابع")
    st.dataframe(resources[['کد منبع', 'نام منبع', 'نوع منبع', 'هزینه واحد (میلیون ریال)']])

    # نمایش تخصیص منابع به فعالیت‌ها
    st.header("تخصیص منابع")
    st.dataframe(assign_res[['کد فعالیت', 'کد منبع', 'نام منبع', 'ساعت برنامه‌ریزی‌شده', 'ساعت واقعی', 'ساعت تفاوت', 'هزینه کل (میلیون ریال)']])

    # نمایش تدارکات
    st.header("تدارکات")
    st.dataframe(procurement[['کد پروژه', 'شماره سفارش خرید', 'کد کالا', 'نام کالا', 'تامین‌کننده', 'تاریخ تحویل', 'وضعیت', 'مبلغ (میلیون ریال)']])

    # نمایش نصب
    st.header("نصب")
    st.dataframe(installation[['کد تجهیز', 'کد پروژه', 'مکان نصب', 'تاریخ نصب', 'درصد پیشرفت', 'وضعیت', 'توضیحات']])
if __name__ == "main":

    main()




