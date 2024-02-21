import os
import pandas as pd
import numpy as np
from datetime import datetime
import re



class StackupFile():

    def __init__(self, filename):
        self.filename = filename
        self.dfs = self.read_excel_file()
            
    def read_excel_file(self):
        try:
            # Attempt to read the Excel file
            dfs = pd.read_excel(self.filename, sheet_name=None, header=None)
            return dfs
        except FileNotFoundError:
            print(f"Error: The file '{self.filename}' was not found.")
        except ValueError as e:
            print(f"Error: There was a problem with the file '{self.filename}'.", e)
        except Exception as e:
            # General exception to catch other potential issues
            print(f"An unexpected error occurred while reading '{self.filename}':", e)
        return dfs
     
    def create_stackup_summary(self, mc_iterations=1000000):
        
        all_results = []
        
        for df in self.dfs.values():
            stackup_class = Stackup(df)
            stackup_results = stackup_class.full_stackup_analysis(mc_iterations)
            all_results.append(stackup_results)
          
        results_df = pd.DataFrame(all_results) 
        
        results_filename = self.create_filename()
        # export dataframe to excel file
        results_df.to_excel(results_filename)
            
        return results_df
        
    def create_filename(self):
    
        # get current date/time
        current_datetime = datetime.now()
        formatted_datetime = current_datetime.strftime('%Y%m%d_%H%M%S')
                
        try:
            # ask for output filename
            output_filename = input('Please enter the desired name of the output file   ')
            self.validate_filename(output_filename)

        except ValueError as e:
            print('uh-oh')
            print(e)
        
        # create filename
        filename = output_filename + '_stackup_report_' + formatted_datetime + '.xlsx'
        
        return filename
        
    def validate_filename(self, filename):
        # Check for invalid characters
        if re.search(r'[\\/:*?"<>|]', filename):
            raise ValueError("Filename contains invalid characters: \\ / : * ? \" < > |")
    
        # Check for length constraints
        if len(filename) > 255:  # General maximum length constraint
            raise ValueError("Filename exceeds the maximum length of 255 characters.")

        return True
        
       
       
class Stackup():

    def __init__(self, df):
        self.stackup_name=df[1][0]
        self.min_goal=df[1][1]
        self.max_goal=df[1][2]
        
        # check for values in stack name and limits
        if pd.isna(self.stackup_name) or pd.isna(self.min_goal) or pd.isna(self.max_goal):
                raise ValueError('Stack name and limit values need to be entered')
        
        # determine what kind of stack limits there are
        if isinstance(self.min_goal, (int, float)):
            if isinstance(self.max_goal, (int, float)):
                self.type = 'both'
            else:
                self.type = 'min_limit'
        elif isinstance(self.max_goal, (int, float)):
            self.type = 'max_limit'
        else:
            print('No numerical limit provided')
           
        # create df
        headers = df.iloc[3]
        self.df = df.iloc[4:].reset_index(drop=True)
        self.df.columns = headers
        
        # drop any rows not needed
        self.df.dropna(thresh=3, inplace=True)
        
        # check to ensure int or floats in columns with data
        columns_to_check = ['Cp', 'Min', 'Max']
        
        for column in columns_to_check:
            for i in self.df[column]:
                if not isinstance(i, (int, float)):
                    raise ValueError(f"Column '{column}' contains non-numeric values.")
        
    def limit_stack(self):
        
        # calculate min and max stack
        min_stack = self.df['Min'].sum()
        max_stack = self.df['Max'].sum()
        
        # create dictionary of results
        results = {'min_limit_stack': min_stack, 'max_limit_stack': max_stack}
        
        return results
        
    def stat_stack(self):
        statdf = self.df
        
        # caclulate sum of squares
        statdf['Range'] = abs(statdf['Min'] - statdf['Max'])
        
        statdf['squares'] = (statdf['Range']/2) ** 2
        
        sumofsquares = (statdf['squares'].sum()) **0.5
        
        
        # calculate 3sigma and 4.5 sigma values
        threesigma = sumofsquares
        fourpointfivesigma = sumofsquares * (4.5/3)
        
        midpoint = np.mean([self.df['Min'].sum(), self.df['Max'].sum()])
        
        negthreesigma = midpoint - threesigma
        posthreesigma = midpoint + threesigma
        negfourpointfivesigma = midpoint - fourpointfivesigma
        posfourpointfivesigma = midpoint + fourpointfivesigma
        
        # create dictionary of results
        results = {'lower_threesigma':negthreesigma,
                   'upper_threesigma':posthreesigma,
                   'lower_fourpointfivesigma':negfourpointfivesigma,
                   'upper_fourpointfivesigma':posfourpointfivesigma}
        
        return results
    
    def monte_carlo(self, iterations=1000000):
        
        mc_raw_data = self.run_monte_carlo(iterations)
        
        statistics = self.monte_carlo_results(mc_raw_data)
        
        return statistics
    
    def run_monte_carlo(self, iterations=1000000):
        mcdf = self.df
        
        # calculate data for monte carlo assuming a normal distribution
        mcdf['midpoints'] = mcdf[['Max', 'Min']].mean(axis=1)
        
        mcdf['stddev'] = abs(mcdf['Max'] - mcdf['Min']) / (6 * mcdf['Cp'])
        
        # run monte carlo
    
        mcdata = []
        for i in mcdf.index:
            nparr = np.random.normal(mcdf.loc[i, 'midpoints'], mcdf.loc[i, 'stddev'], iterations)
            mcdata.append(nparr)
        
        mc_simulation_results = np.add.reduce(np.array(mcdata))
            
        return mc_simulation_results
    
    def monte_carlo_results(self, data):
            
        # calculate statistics for monte carlo
            
        mean = data.mean()
        min_ = data.min()
        max_ = data.max()
        median = np.median(data)
        
        
        # determine NOK results
        if self.type == 'both':
            less_than = np.sum(data < self.min_goal)
            greater_than = np.sum(data > self.max_goal)
            percent_less_than = less_than / len(data)
            percent_greater_than = greater_than/len(data)
                
        elif self.type == 'min_limit':
            less_than = np.sum(data < self.min_goal)
            percent_less_than = less_than / len(data)
            percent_greater_than = 0
            
        elif self.type == 'max_limit':
            greater_than = np.sum(data > self.max_goal)
            percent_greater_than = greater_than/len(data)
            percent_less_than = 0
                
        percent_NOK = percent_less_than + percent_greater_than
            
        # dictionary of results
        mc_results_dict = {'mc_mean': mean, 'mc_min': min_, 'mc_max': max_, 
                           'mc_median': median,
                           'mc_percent_less': percent_less_than,
                           'mc_percent_more': percent_greater_than,
                           'mc_percent_NOK': percent_NOK}    
            
        return mc_results_dict
         
    def full_stackup_analysis(self,iterations=1000000):
            
        # run all stack ups and return results in a dict
        
        stack_name = {'name':self.stackup_name}
        
        goals = {'min_goal': self.min_goal, 'max_goal':self.max_goal}
        
        limit_stack_results = self.limit_stack()
            
        stat_stack_results = self.stat_stack()
            
        mc_results = self.monte_carlo(iterations)
            
        all_results = {**stack_name, **goals, **limit_stack_results, **stat_stack_results, **mc_results}
            
        return all_results
            

            
input_filename = input('Please enter the filename and include the .xlsx extension   ')

stackups = StackupFile(input_filename)

results = stackups.create_stackup_summary()



