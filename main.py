import pandas as pd
from scripts.dataEngine import ThermalDispatchInput  # Correct import
from scripts.dispatchBase import DispatchModel
from scripts.dispatchPpa import DispatchModelPPA
from scripts.report import OutputReport
from datetime import datetime
import time

# Defining the file path for the Excel file
file_path = "inputs/Thermal_Dispatch_Input.xlsx"

# Creating an instance of the ThermalDispatchInput class
time_1 = time.time()
print("Reading Inputs...")
thermal_input = ThermalDispatchInput(file_path)
time_2 = time.time()
print(f"Reading Inputs finished (took {time_2 - time_1} seconds)...")
# Creating a new dataframe based on df_lmp
df_new = thermal_input.df_lmp.copy()

time_3 = time.time()
print("Getting inputs for base...")
# Apply functions in order
df_new = thermal_input.add_gas(df_new)
df_new = thermal_input.add_vom(df_new)
df_new = thermal_input.add_gox(df_new, "NOx")
df_new = thermal_input.add_gox(df_new, "SOx")
df_new = thermal_input.add_gox(df_new, "CO2")

# Base model inputs
gas_price_col = thermal_input.get_hub("Gas_Hub","Base")
power_price_col = thermal_input.get_hub("LMP_Hub","Base")
nox_zone = thermal_input.get_gox_zone("NOx", "Base")
co2_zone = thermal_input.get_gox_zone("CO2", "Base")
sox_zone = thermal_input.get_gox_zone("SOx", "Base")
t_lowers, t_uppers = thermal_input.get_start_time("Base")
heat_rate = thermal_input.get_heat_rate("Base")
MinUpTime = 8
MinDownTime = 8
Startcost_hot, Startcost_warm, Startcost_cold = thermal_input.get_start_cost("Base")
sox_rate, nox_rate, co2_rate = thermal_input.get_gox_rate(name1 = "SOx_Rate", name2 = "NOx_Rate", name3="CO2_Rate")
mincap, maxcap = thermal_input.get_cap("Base")
vom_type = thermal_input.get_vom_type("Base")
status = thermal_input.check_ppa_status("Base")
ltsa = thermal_input.get_ltsa("Base")
eoh = thermal_input.get_eoh("Base")
df_new["Adder_st"] = df_new.apply(thermal_input.get_adder, axis=1)
df_new["Adder_bid"] = df_new.apply(thermal_input.get_adder_bid, axis=1)
df_new.drop(columns=["Weekday", "Type", "Leap"], axis=1, inplace=True)
time_4 = time.time()
print(f"Input to base finished (took {time_4-time_3} seconds)")

if status == True:
    st_year, end_year = thermal_input.get_time("Base")
    data = thermal_input.get_data_file("Base", df_new)
    results_list = [] 
    for year in range(st_year, end_year+1):   #st_year is 2025
        df_new_base = thermal_input.get_year_data(data, year)
        maint_per = thermal_input.get_maint_per(year)
        if df_new_base.shape[0] ==0:
            print(f"Null year {year}...")
            print("*************************")
            continue
        else:
            T = df_new_base.shape[0]
            print(f"Running year {year}...")
            dispatchModelBase = DispatchModel(df_new_base, gas_price_col, power_price_col, nox_zone, co2_zone, sox_zone, 
                 t_lowers, t_uppers, heat_rate, MinUpTime, MinDownTime, Startcost_hot, 
                 Startcost_warm, Startcost_cold, sox_rate, nox_rate, co2_rate, 
                 mincap, maxcap, vom_type, maint_per, T, ltsa, eoh)
    
            print("Starting Solver...")
            dispatchModelBase.solve()
            print("Solver Stoped")
            print("**********************")
            dispatchModelBase.get_results()
            yearly_result = dispatchModelBase.get_results()
            results_list.append(yearly_result)

# Concatenate all yearly results into one DataFrame
final_results = pd.concat(results_list, ignore_index=True)

    
print("Base Run Sucessful")
print("**************************")
n = thermal_input.df_ppa.shape[1]
for i in range(2,n):
    name = thermal_input.df_ppa.columns[i]
    status = thermal_input.check_ppa_status(name)
    if status == False:
        continue
    else:
        st_year, end_year = thermal_input.get_time(name)
        gas_price_col = thermal_input.get_hub("Gas_Hub",name)
        power_price_col = thermal_input.get_hub("LMP_Hub",name)
        nox_zone = thermal_input.get_gox_zone("NOx", name)
        co2_zone = thermal_input.get_gox_zone("CO2", name)
        sox_zone = thermal_input.get_gox_zone("SOx", name)
        t_lowers, t_uppers = thermal_input.get_start_time(name)
        heat_rate = thermal_input.get_heat_rate(name)
        Startcost_hot, Startcost_warm, Startcost_cold = thermal_input.get_start_cost(name)
        mincap, maxcap = thermal_input.get_cap(name)
        vom_type = thermal_input.get_vom_type(name)
        data = thermal_input.get_data_file(name, final_results)
        ltsa = thermal_input.get_ltsa(name)
        eoh = thermal_input.get_eoh(name)
        mover_dep = thermal_input.get_mover_dependency(name)
        print(f"mover dependency {mover_dep}")
        if pd.notna(mover_dep):
            print(f"running for {name}...")
            results_list = [] 
            for year in range(st_year, end_year+1):   #st_year is 2025
                df_new_ppa = thermal_input.get_year_data(data, year)
                if df_new_ppa.shape[0] ==0:
                    print(f"Null year {year}...")
                    print("**********************")
                    continue
                else:
                    T = df_new_ppa.shape[0]
                    print(f"Running year {year}...")
                    dispatchmodelppa = DispatchModelPPA(df_new_ppa, gas_price_col, power_price_col, nox_zone, 
                                                        co2_zone, sox_zone, t_lowers, t_uppers, heat_rate, 
                                                        Startcost_hot, Startcost_warm, Startcost_cold, sox_rate, nox_rate, co2_rate, mincap, 
                                                        maxcap, vom_type,mover_dep,name, T, ltsa, eoh)

                    print(f"Starting Solver...")
                    dispatchmodelppa.solve()
                    print(f"Solver Stoped")
                    print("**********************")
                    dispatchModelBase.get_results()
                    yearly_result = dispatchmodelppa.get_results()
                    results_list.append(yearly_result)

            final_results_ppa = pd.concat(results_list, ignore_index=True)

            if final_results.shape[0] == final_results_ppa.shape[0]:
                final_results = final_results_ppa
            else:
                combined_df = pd.concat([final_results, final_results_ppa], axis=0, ignore_index=True)
                combined_df['ON_'+name] = combined_df['ON_'+name].fillna(0)
                combined_df['Power_'+name] = combined_df['Power_ppa'+name].fillna(0)

        else:
            print(f"running for {name}...")
            results_list = [] 
            for year in range(st_year, end_year+1): 
                df_new_ppa = thermal_input.get_year_data(data, year)
                if df_new_ppa.shape[0] ==0:
                    print(f"Null year {year}...")
                    print("**********************")
                    continue
                else:
                    T = df_new_ppa.shape[0]
                    maint_per = thermal_input.get_maint_per(year)
                    print(f"Running year {year}...")
                    dispatchModelBase = DispatchModel(df_new_ppa, gas_price_col, power_price_col, nox_zone,
                                            co2_zone, sox_zone, t_lowers, t_uppers, heat_rate,
                                            MinUpTime, MinDownTime, Startcost_hot, Startcost_warm, 
                                            Startcost_cold, sox_rate, nox_rate, co2_rate, mincap, 
                                            maxcap, vom_type, maint_per,T,ltsa, eoh, name)

                    print(f"Starting Solver...")
                    dispatchModelBase.solve()
                    print(f"Solver Stoped")
                    print("**********************")
                    dispatchModelBase.get_results()
                    yearly_result = dispatchModelBase.get_results()
                    results_list.append(yearly_result)

            final_results_ppa = pd.concat(results_list, ignore_index=True)

            if final_results.shape[0] == final_results_ppa.shape[0]:
                final_results = final_results_ppa
            else:
                merge_keys = ["Year", "Month", "Day", "Hour"]

                # Merge PPA values based on date/time only
                final_results = final_results.merge(
                                final_results_ppa[merge_keys + ["ON_"+name, "Power_"+name]],
                                on=merge_keys,
                                how="left"
                                )

                # Fill missing values with 0
                final_results["ON_"+name] = final_results["ON_"+name].fillna(0)
                final_results["Power_"+name] = final_results["Power_"+name].fillna(0)


timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
filename1 = f"output/ON_base_{timestamp}.xlsx"
filename2 = f"output/Percent On_base_{timestamp}.xlsx"
filename3 = f"output/dispatch_results_{timestamp}.xlsx"

output_base = OutputReport(final_results, 2025, 2050)
test_file1 = output_base.get_output("PPA2")
test_file2 = output_base.compute_monthly_percentage()

test_file1.to_excel(filename1, index = True)
test_file2.to_excel(filename2, index = True)
final_results.to_excel(filename3, index=False)

for i in range(1,thermal_input.df_ppa.shape[1]):
    ppa_name = thermal_input.df_ppa.columns[i]
    status = thermal_input.check_ppa_status(ppa_name)
    if status == True:
        _,maxcap = thermal_input.get_cap(ppa_name)
        df1 = output_base.get_capacity_factor(ppa_name, final_results, maxcap)
        df2 = output_base.get_pivot_table(final_results, ppa_name)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        df1.to_excel(f"output/CF_{ppa_name}_{timestamp}.xlsx", index = False)
        df2.to_excel(f"output/Pivot Table_{ppa_name}_{timestamp}.xlsx", index = False)
    

time_finish = time.time()
print(f"Time taken to finish the process {time_finish - time_1} seconds")
print("Thank you for using the toolkit")