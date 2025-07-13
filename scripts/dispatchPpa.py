from pyomo.environ import *
import pandas as pd
import numpy as np

class DispatchModelPPA:
    def __init__(self, data, gas_price_col, power_price_col, nox_zone, co2_zone, sox_zone, 
                 t_lowers, t_uppers, heat_rate, Startcost_hot, Startcost_warm, Startcost_cold, 
                 sox_rate, nox_rate, co2_rate, mincap, maxcap, vom_type, mover_dep, name, T, ltsa, eoh):
        
        self.data = data.copy()
        self.T = T
        self.Startcost_hot, self.Startcost_warm, self.Startcost_cold = Startcost_hot, Startcost_warm, Startcost_cold
        self.mincap, self.maxcap = mincap, maxcap
        self.ltsa = ltsa
        self.eoh = eoh
        self.t_lowers, self.t_uppers = t_lowers, t_uppers
        self.heat_rate = heat_rate
        self.sox_rate, self.nox_rate, self.co2_rate = sox_rate, nox_rate, co2_rate
        self.mover_dep = mover_dep
        self.name = name
        
        # Vectorized Cost Calculation
        self.cost = self.heat_rate * (self.data[gas_price_col] / 1000).reset_index(drop=True)
        self.price = self.data[power_price_col].reset_index(drop=True)
        self.gas_price = self.data[gas_price_col].reset_index(drop=True)
        self.co2_cost = (self.data[co2_zone] * co2_rate).fillna(0).reset_index(drop=True)
        self.sox_cost = (self.data[sox_zone] * sox_rate).fillna(0).reset_index(drop=True)
        self.nox_cost = (self.data[nox_zone] * nox_rate).fillna(0).reset_index(drop=True)
        self.varCost = self.data[vom_type].reset_index(drop=True)
        self.st_adder = self.data["Adder_st"].reset_index(drop=True)
        self.bidAdder = self.data["Adder_bid"].reset_index(drop=True)
        self.mover_on = self.data["ON_"+ mover_dep].reset_index(drop=True)

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
        
        total_emission_cost = self.co2_cost + self.sox_cost + self.nox_cost
        
        m.obj = Objective(
            expr=sum(self.price[t]* m.elect[t] - (self.cost[t]* m.elect[t] 
                     + self.varCost[t]* m.elect[t])*(1+self.bidAdder[t]) - total_emission_cost[t]*m.elect[t] 
                     - (self.Startcost_hot) * m.switch_on[t]  - self.eoh*m.ON[t] 
                     -m.switch_on[t]*(self.st_adder[t]+self.ltsa) for t in range(self.T)),
            sense=maximize
        )
        self._define_constraints()

    def _define_constraints(self):
        m = self.model
        
        # Switch Constraints
        m.switch_constraint = ConstraintList()
        for t in range(1, self.T): 
            m.switch_constraint.add(m.switch_on[t] >= m.ON[t] - m.ON[t - 1])
            m.switch_constraint.add(m.switch_off[t] >= m.ON[t-1] - m.ON[t])
            m.switch_constraint.add(m.switch_off[t] +  m.ON[t] == m.ON[t - 1] + m.switch_on[t])
        
        #mover dependecy constraint
        m.mover_dependency = ConstraintList()
        for t in range(self.T):
            m.mover_dependency.add(m.ON[t] <= self.mover_on[t])


        #Start type constraint
        if self.t_lowers != 0 and self.t_uppers !=0:
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

        # Capacity Constraints
        m.cap = ConstraintList()
        for t in range(self.T):    
            m.cap.add(m.elect[t] <= self.maxcap * m.ON[t])
            m.cap.add(m.elect[t] >= self.mincap * m.ON[t])

    def solve(self, solver_name='scip', solver_path='scip.exe'):
        solver = SolverFactory(solver_name, executable=solver_path)
        solver.solve(self.model)

    def get_results(self):
        # Collect the results
        results = {
            "ON": [self.model.ON[t].value for t in range(self.T)],
            "Power": [self.model.elect[t].value for t in range(self.T)]
        }

        # Add the results to the existing DataFrame
        self.data.loc[:,"ON_" + self.name] = results["ON"]
        self.data.loc[:,"Power_"+ self.name] = results["Power"]

        # Save the updated DataFrame to a new Excel file
        return self.data


