import pandas as pd
import numpy as np
from calendar import monthrange
from datetime import datetime

class OutputReport:
    def __init__(self, data, start_year, end_year):

        self.data = data
        self.start_year = start_year
        self.end_year = end_year
        self._get_dict_ready_()


    def _get_dict_ready_(self):
        self.years = []
        for i in range(self.start_year, self.end_year + 1):
            self.years.append(i)
        
        n = len(self.years)
        data_out = {
            "Jan": [0] * n,
            "Feb": [0] * n,
            "Mar": [0] * n,
            "Apr": [0] * n,
            "May": [0] * n,
            "Jun": [0] * n,
            "Jul": [0] * n,
            "Aug": [0] * n,
            "Sep": [0] * n,
            "Oct": [0] * n,
            "Nov": [0] * n,
            "Dec": [0] * n,
        }

        df = pd.DataFrame(data_out, index=self.years).T
        df.index.name = "Month"
        self.data_out = df

    
    def get_output(self, ppa_type):
        month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        
        self.month_data = self.data_out.copy()

        for t in range(len(self.data)):
            yr = self.data.loc[t, "Year"]
            if yr in self.years:
                month_raw = self.data.loc[t, "Month"]
                if pd.notna(month_raw):
                    month_num = int(month_raw)
                    is_on = self.data.loc[t, f"ON_{ppa_type}"]
                    if is_on == 1 and 1 <= month_num <= 12:
                        month_str = month_names[month_num - 1]
                        self.data_out.at[month_str,yr] += 1
                    if 1<= month_num <= 12:
                        month_str = month_names[month_num - 1]
                        self.month_data.at[month_str,yr] += 1
    
        return self.data_out 

    def compute_monthly_percentage(self):
        """Compute percentage contribution of each month to total ON count in a year"""
        df = self.data_out.copy()
        df1 = self.month_data.copy()
            
        df_percent = (df/df1).fillna(0) * 100 
        self.data_out_percent = df_percent

        return self.data_out_percent
    
    def get_capacity_factor(self, ppa_name, result, maxcap):
        results_group = result.groupby(["Year","Month"])["Power_"+ppa_name].sum().reset_index()
        results_group["Power_"+ppa_name] = results_group.apply(
            lambda row: row["Power_"+ppa_name] / (maxcap*24 * monthrange(row["Year"].astype(int), row["Month"].astype(int))[1]),
            axis=1
        )

        return results_group

    def get_pivot_table(self, result, ppa_name):
        results_pivot = round(pd.pivot_table(result, 
                                     values="Power_"+ppa_name,
                                     index='Month', 
                                     columns='Year', 
                                     aggfunc="sum"),2) 
        return results_pivot       




    