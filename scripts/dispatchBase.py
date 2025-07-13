from pyomo.environ import *
import pandas as pd
import numpy as np
import math

class DispatchModel:
    def __init__(self, data, gas_price_col, power_price_col, nox_zone, co2_zone, sox_zone, 
                 t_lowers, t_uppers, heat_rate, MinUpTime, MinDownTime, Startcost_hot, 
                 Startcost_warm, Startcost_cold, sox_rate, nox_rate, co2_rate, 
                 mincap, maxcap, vom_type, maint_per, T, ltsa, eoh, name = "Base", check_maint_con=False):
        
        self.data = data.copy()
        self.T = T
        self.MinUpTime = MinUpTime
        self.MinDownTime = MinDownTime
        self.Startcost_hot, self.Startcost_warm, self.Startcost_cold = Startcost_hot, Startcost_warm, Startcost_cold
        self.mincap, self.maxcap = mincap, maxcap
        self.t_lowers, self.t_uppers = t_lowers, t_uppers
        self.heat_rate = heat_rate
        self.sox_rate, self.nox_rate, self.co2_rate = sox_rate, nox_rate, co2_rate
        self.maint_per = maint_per
        self.name = name
        self.ltsa = ltsa
        self.eoh = eoh
        self.check_maint_con = check_maint_con
        self.month_days = [744, 672, 744, 720, 744, 720, 744, 744, 720, 744, 720, 744]
        self.start_index = [0, 744, 1416, 2160, 2880, 3624, 4344, 5088, 5832, 6552, 7296, 8016]
        
        self.cost = self.heat_rate * (self.data[gas_price_col] / 1000).reset_index(drop=True)
        self.price = self.data[power_price_col].reset_index(drop=True)
        self.gas_price = self.data[gas_price_col].reset_index(drop=True)
        self.co2_cost = (self.data[co2_zone] * co2_rate).fillna(0).reset_index(drop=True)
        self.sox_cost = (self.data[sox_zone] * sox_rate).fillna(0).reset_index(drop=True)
        self.nox_cost = (self.data[nox_zone] * nox_rate).fillna(0).reset_index(drop=True)
        self.varCost = self.data[vom_type].reset_index(drop=True)
        self.bidAdder = self.data["Adder_bid"].reset_index(drop=True)
        self.st_adder = self.data["Adder_st"].reset_index(drop=True)

        self.model = ConcreteModel()
        self._build_model()

    def init_a_ij_rule(model, i, j):
            """Initializes delta_type variables to 0."""
            return 0
    
    def _build_model(self):
        def init_a_ij_rule(model, i, j):
            """Initializes delta_type variables to 0."""
            return 0
        m = self.model
        m.I = Set(initialize=range(self.T))
        m.J = Set(initialize=range(3))
        m.ON = Var(range(self.T), domain=Binary)
        m.switch_on = Var(range(self.T), domain=Binary)
        m.switch_off = Var(range(self.T), domain = Binary)
        m.delta_type = Var(m.I, m.J, domain= Binary, initialize=init_a_ij_rule)
        m.start_cost = Var(range(self.T), domain=NonNegativeReals)
        m.elect = Var(range(self.T), domain=NonNegativeReals)
        if self.check_maint_con == True:
            artvar_index = []
            for i in range(len(self.maint_per)):
                if self.maint_per != 0.0:
                    k = math.ceil(self.month_days[i] * self.maint_per[i])
                    for t in range(self.month_days[i] - k):
                        artvar_index.append((i, t))

            m.artvar_index = Set(initialize=artvar_index, dimen=2)
            m.artvar = Var(m.artvar_index, domain=Binary)
        
        total_emission_cost = self.co2_cost + self.sox_cost + self.nox_cost
        
        m.obj = Objective(
            expr=sum(self.price[t]* m.elect[t] - (self.cost[t]* m.elect[t] 
                     + self.varCost[t]* m.elect[t])*(1+self.bidAdder[t])- total_emission_cost[t]*m.elect[t] 
                     - ((self.Startcost_hot) * m.delta_type[t,0] + (self.Startcost_warm) * m.delta_type[t,1] 
                        + (self.Startcost_cold) * m.delta_type[t,2]) - self.eoh*m.ON[t] 
                     -m.switch_on[t]*(self.st_adder[t] + self.ltsa) for t in range(self.T)),
            sense=maximize
        )
        self._define_constraints()

    def _define_constraints(self):
        m = self.model

        # Up-time
        m.up_time = ConstraintList()
        for t in range(self.MinUpTime-1,self.T):
            m.up_time.add(sum(m.switch_on[i] for i in range(t-self.MinUpTime+1,t))<=m.ON[t])
        
        # Down-time Constraints
        m.down_time = ConstraintList()
        for t in range(self.MinDownTime-1,self.T):
            m.down_time.add(sum(m.switch_off[i] for i in range(t-self.MinDownTime+1,t))<=1-m.ON[t])


        #Start type constraint
        m.delta_start = ConstraintList()
        for i in range(2):
            if i==0:
                p = 1
                q = self.t_lowers
                t = self.t_lowers
            else:
                p = self.t_lowers
                q = self.t_uppers
                t = self.t_uppers
            for j in range(t,self.T):
                m.delta_start.add(m.delta_type[j,i] <= sum(m.switch_off[j-z] for z in range(p,q)))
        
        # Switch Constraints
        m.switch_constraint = ConstraintList()
        for t in range(1, self.T): 
            m.switch_constraint.add(m.switch_on[t] >= m.ON[t] - m.ON[t - 1])
            m.switch_constraint.add(m.switch_off[t] >= m.ON[t-1] - m.ON[t])
            m.switch_constraint.add(m.switch_off[t] +  m.ON[t] == m.ON[t - 1] + m.switch_on[t])

        # #Start type selection at switch on
        m.delta_sum = ConstraintList() 
        for t in range(self.T):
            m.delta_sum.add(sum(m.delta_type[t,i] for i in range(3)) >= m.switch_on[t])

        #Maintenance Constraint
        if self.check_maint_con == True:
            m.maint_cons = ConstraintList()
            m.maint = ConstraintList()
            for i in range(len(self.maint_per)):
                if self.maint_per[i]==0.0:
                    continue
                else:
                    st = self.start_index[i]
                    k = math.ceil(self.month_days[i]*self.maint_per[i])
                    for t in range(self.month_days[i]-k):
                        m.maint.add(m.switch_off[st+t] + m.switch_on[st+k+t] >= 2*m.artvar[i,t])
                        m.maint.add(sum(m.switch_off[st+i+t] for i in range(k+1)) <= (k+1)*(1-m.artvar[i,t])+1)
                        m.maint.add(sum(m.switch_on[st+i+t] for i in range(k+1)) <= (k+1)*(1-m.artvar[i,t])+1)
        
                    m.maint_cons.add(sum(m.artvar[i,t] for t in range(self.month_days[i]-k))>=1)

        else:
            m.maint_cons = ConstraintList()
            for i in range(len(self.maint_per)):
                if self.maint_per[i]==0.0:
                    continue
                else:
                    st = self.start_index[i]
                    k = math.ceil(self.month_days[i]*self.maint_per[i])
                    m.maint_cons.add(sum(m.ON[j] for j in range(st, st + self.month_days[i]+1))<= self.month_days[i]-k)


        # Capacity Constraints
        m.cap = ConstraintList()
        for t in range(self.T):    
            m.cap.add(m.elect[t] <= self.maxcap * m.ON[t])
            m.cap.add(m.elect[t] >= self.mincap * m.ON[t])

    def solve(self, solver_name='scip', solver_path='scip.exe'):
        solver = SolverFactory(solver_name, executable=solver_path)
        solver.options['limits/gap'] = 0.003
        solver.solve(self.model) #timelimit=300,

    def get_results(self):
        # Collect the results
        results = {
            "ON": [self.model.ON[t].value for t in range(self.T)],
            "Power": [self.model.elect[t].value for t in range(self.T)]
        }

        # Add the results to the existing DataFrame
        # self.data["ON_"+ self.name] = results["ON"]
        # self.data["Power_"+ self.name] = results["Electricity Production"]
        self.data.loc[:, "ON_" + self.name] = results["ON"]
        self.data.loc[:, "Power_" + self.name] = results["Power"]

        # Save the updated DataFrame to a new Excel file
        return self.data


