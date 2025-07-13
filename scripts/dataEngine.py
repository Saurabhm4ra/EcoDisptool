import pandas as pd
from datetime import datetime

class ThermalDispatchInput:
    def __init__(self, file_path):
        self.file_path = file_path
        self._read_inputs()
        self.month_mapping = {
            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun', 
            7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
            }
        

    def _read_inputs(self):
        """ Reads all the input data from the excel file """
        with pd.ExcelFile(self.file_path) as xlsModel:
            # Read tabs of the input sheet
            self.df_lmp = pd.read_excel(xlsModel, 'LMP')
            self.df_param = pd.read_excel(xlsModel, 'Plant_Param')
            self.df_gas = pd.read_excel(xlsModel, 'Gas')
            self.df_nox = pd.read_excel(xlsModel, 'NOx')
            self.df_sox = pd.read_excel(xlsModel, 'SOx')
            self.df_co2 = pd.read_excel(xlsModel, 'CO2')
            self.df_vom = pd.read_excel(xlsModel, 'VOM')
            self.df_maint = pd.read_excel(xlsModel, 'Maint')
            self.df_ppa = pd.read_excel(xlsModel, 'PPA')
            self.df_adder_st = pd.read_excel(xlsModel, 'Adder ($ per start)')
            self.df_adder_bid = pd.read_excel(xlsModel, 'Adder(% Bid)')
    
    def add_gox(self, data, name):
        if name == "NOx":
            for i in range(2, self.df_nox.shape[1]):  # Start from 2nd column
                zone_name = self.df_nox.columns[i]  # Get column name
                new_name = f"{name}_{zone_name}"  # Construct new column name
                data = data.merge(self.df_sox[["Year", "Month", zone_name]], on=["Year", "Month"], how="left")
                data.rename(columns={zone_name: new_name}, inplace=True)
                data[new_name] = data[new_name].fillna(0)

        elif name == "SOx":
            for i in range(2, self.df_sox.shape[1]):
                zone_name = self.df_sox.columns[i]
                new_name = f"{name}_{zone_name}"
                data = data.merge(self.df_sox[["Year", "Month", zone_name]], on=["Year", "Month"], how="left")
                data.rename(columns={zone_name: new_name}, inplace=True)
                data[new_name] = data[new_name].fillna(0)  

        elif name == "CO2":
            for i in range(2, self.df_co2.shape[1]):
                zone_name = self.df_co2.columns[i]
                new_name = f"{name}_{zone_name}"
                data = data.merge(self.df_co2[["Year", "Month", zone_name]], on=["Year", "Month"], how="left")
                data.rename(columns={zone_name: new_name}, inplace=True)
                data[new_name] = data[new_name].fillna(0)
  

        return data
    
    def add_vom(self, data, name="VOM"):
        type_mapping = {}
        data_ppa = self.df_ppa.set_index("Param")
        vom_cols = []  # Keep track of new columns added

        for i in range(data_ppa.shape[1]):
            vom_type = data_ppa.loc[name, data_ppa.columns[i]]
            if vom_type not in type_mapping:
                type_mapping[vom_type] = self.df_vom[["Year", "Month", vom_type]].copy()

            new_col = data_ppa.columns[i] + "_" + vom_type
            vom_cols.append(new_col)

            data = data.merge(
                type_mapping[vom_type].rename(columns={vom_type: new_col}),
                on=["Year", "Month"], how="left"
            )

        # Fill NaNs in only the VOM columns with 0
        data[vom_cols] = data[vom_cols].fillna(0)

        return data

    def add_gas(self, data):
        """ Add gas data to the input data """
        for i in range(2, self.df_gas.shape[1]):
            data = data.merge(self.df_gas[["Year", "Month", self.df_gas.columns[i]]], on=["Year", "Month"], how="left")

        return data

    def get_start_time(self, col_name):
        """ Get the hot start and cold start duration based on the selected column name """
        data_ppa = self.df_ppa.set_index("Param")
        # hot_st = int(data_ppa.loc["Hot Start Duration (hour)", col_name][1:] or 0)
        val_hot = data_ppa.loc["Hot Start Duration (hour)", col_name]
        if isinstance(val_hot, str) and len(val_hot) > 1:
            hot_st = int(val_hot[1:])
        else:
            hot_st = 0

        # cold_st = int(data_ppa.loc["Cold Start Duration (hour)", col_name][1:] or 0)
        val_cold = data_ppa.loc["Cold Start Duration (hour)", col_name]
        if isinstance(val_cold, str) and len(val_cold) > 1:
            cold_st = int(val_cold[1:])
        else:
            cold_st = 0
        return hot_st, cold_st

    def get_start_cost(self, col_name):
        """ Get the start costs (Hot, Warm, Cold) based on the selected column name """
        data_ppa = self.df_ppa.set_index("Param")
    
        def get_cost(row_label):
            val = data_ppa.loc[row_label, col_name]
            return float(val) if pd.notna(val) else 0

        hot_st_cost = get_cost("Hot Start Cost ($/start)")
        warm_st_cost = get_cost("Warm start Cost ($/start)")
        cold_st_cost = get_cost("Cold start Cost ($/start)")
    
        return hot_st_cost, warm_st_cost, cold_st_cost


    def get_ramp_rate(self, col_name):
        """ Get the ramp up rate based on the selected column name """
        data_ppa = self.df_ppa.set_index("Param")
        ramp = float(data_ppa.loc["Ramp up rate (MW/min)", col_name] or 0)

        return ramp

    def get_heat_rate(self, col_name):
        """ Get the heat rate based on the selected column name """
        data_ppa = self.df_ppa.set_index("Param")
        h_rate = float(data_ppa.loc["Heat_Rate", col_name] or 0)

        return h_rate

    def get_data_file(self, col_name, data_inp):
        """ Filter the LMP data based on the PPA start and end date """
        data = self.df_ppa.set_index("Param")
        start_date = data.loc["PPA_Start", col_name]
        end_date = data.loc["PPA_End", col_name]

        filtered_df = data_inp[(data_inp["Date"] >= start_date) & (data_inp["Date"] <= end_date)]

        return filtered_df
    
    def get_year_data(self, data, year):
        filtered_df = data[data["Year"] == year]
        return filtered_df
    
    def get_num_days(self, data, month):
        month_dict = {"Jan": 1,
                      "Feb": 2,
                      "Mar": 3,
                      "Apr": 4,
                      "May": 5,
                      "Jun": 6,
                      "Jul": 7,
                      "Aug": 8,
                      "Sep": 9,
                      "Oct": 10,
                      "Nov": 11,
                      "Dec": 12}
        value = month_dict.get(month)
        return (data["Month"]== value).sum()
    
    def get_gox_rate(self, name1, name2, name3):
        data = self.df_param.set_index("Param")
        return float(data.loc[name1, "Value"]), float(data.loc[name2, "Value"]), float(data.loc[name3, "Value"]) 

    def get_hub(self, name_hub, name_ppa):
        data = self.df_ppa.set_index("Param")
        return data.loc[name_hub, name_ppa]

    def get_gox_zone(self, name_gas, name_ppa):
        data = self.df_ppa.set_index("Param")
        return name_gas + "_" + data.loc[name_gas, name_ppa]
    
    def get_cap(self, name_ppa):
        data = self.df_ppa.set_index("Param")
        return float(data.loc["Contracted_Min", name_ppa]),float(data.loc["Contracted_Cap", name_ppa]) 
    
    def get_vom_type(self, name_ppa):
        data = self.df_ppa.set_index("Param")
        return name_ppa + "_" + data.loc["VOM", name_ppa]
    
    def check_ppa_status(self, name_ppa):
        data = self.df_ppa.set_index("Param")
        return data.loc["Dispatch_Run", name_ppa]

    def get_time(self, col_name):
        data = self.df_ppa.set_index("Param")
        start_date = data.loc["PPA_Start", col_name]
        end_date = data.loc["PPA_End", col_name]
        st_date_obj = datetime.strptime(str(start_date), "%Y-%m-%d %H:%M:%S")
        end_date_obj = datetime.strptime(str(end_date), "%Y-%m-%d %H:%M:%S")
        st_year = st_date_obj.year
        end_year = end_date_obj.year
        return st_year, end_year
    
    def get_adder(self, row):
        year = row['Year']
        month = row['Month']
        month_name = self.month_mapping.get(month, None)
    
        if month_name and year in self.df_adder_st.index:
            return self.df_adder_st.at[year, month_name]
        return 0  # Default to zero if no match found

    def get_adder_bid(self, row):
        year = row['Year']
        month = row['Month']
        month_name = self.month_mapping.get(month, None)
    
        if month_name and year in self.df_adder_bid.index:
            return self.df_adder_bid.at[year, month_name]
        return 0  # Default to zero if no match found
        
    
    def get_mover_dependency(self, name):
        data = self.df_ppa.set_index("Param")
        return data.loc["Mover_Dependency", name]

    def get_maint_per(self, year):
        data = self.df_maint.set_index("Year")
        months = data.columns

        if year in data.index:
            row = data.loc[year]
            maint = []
            for val in row:
                if pd.notnull(val):  # Check if val exists (not NaN or None)
                    maint.append(val) 
                else:
                    maint.append(0.0)
        else:
            maint = [0.0] * len(months)

        return maint
    
    def get_eoh(self, col_name):
        data = self.df_ppa.set_index("Param")
        eoh = data.loc["EOH ($/h)", col_name]

        if pd.notna(eoh):
            return eoh
        else:
            return 0
        
    def get_ltsa(self, col_name):
        data = self.df_ppa.set_index("Param")
        ltsa = data.loc["LTSA ($/start)", col_name]

        if pd.notna(ltsa):
            return ltsa
        else:
            return 0










