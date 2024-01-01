import collections
import matplotlib.pyplot as plt
import os 
from ortools.sat.python import cp_model
import DeeplearningModel
import numpy as np
class MESJobShop:
    def __init__(self):
        self.order_add = []
        self.ScheduleData=[] #task_data(machine, oPos, start_time, duration)
        self.P2P_time = [75, 17, 12, 27, 20, 27, 24, 13, 27, 42, 14, 50]  # 輸送帶運送時間
        self.GattPlot = True
        self.language = 'EN' # 甘特圖語言設定 'TW' for Tranditional Chinese, 'EN' for English
        self.output = ""
        self.__model = cp_model.CpModel()
        self.__bar_color = ['red', 'blue', 'green', 'orangered', 'gold', 'purple', 'teal', 'limegreen', 'navy', 'orange',] # 甘特圖工單代表顏色
        self.__machine_name_en = ["Storage", 'AGV', 'Check part', 'Milling', 'PCB', 'Fuse check', 'Cover', 'Pressure', 'Heat']
        self.__machine_name_tw = ['倉儲　　　', '無人搬運車', '影像檢測　', '鑽孔　　　', 'PCB放置 　', '保險絲檢測', '上蓋放置　', '加壓　　　', '加熱　　　']
    def __del__(self):
        self.__model = None
    def CPmodelSovel(self): 
        task_data = collections.namedtuple('task_data', 'machine, oPos, start_time, duration')
        machines_count = 1 + max(task[0] for job in self.jobs_data for task in job)
        all_machines = range(machines_count)
        # Computes horizon dynamically as the sum of all durations.
        horizon = sum(task[1] for job in self.jobs_data for task in job) + 308*20
        # Named tuple to store information about created variables.
        task_type = collections.namedtuple('task_type', 'start end interval')
        # Named tuple to manipulate solution information.
        assigned_task_type = collections.namedtuple('assigned_task_type','start job index duration')
        # Creates job intervals and add to the corresponding machine lists.
        all_tasks = {}
        machine_to_intervals = collections.defaultdict(list)
        for job_id, job in enumerate(self.jobs_data):
            for task_id, task in enumerate(job):
                machine = task[0]
                duration = task[1]
                suffix = '_%i_%i' % (job_id, task_id)
                start_var = self.__model.NewIntVar(0, horizon, 'start' + suffix)
                end_var = self.__model.NewIntVar(0, horizon, 'end' + suffix)
                interval_var = self.__model.NewIntervalVar(start_var, duration, end_var,
                                                    'interval' + suffix)
                all_tasks[job_id, task_id] = task_type(start=start_var,
                                                    end=end_var,
                                                    interval=interval_var)
                machine_to_intervals[machine].append(interval_var)
        # Create and add disjunctive constraints.
        for machine in all_machines:
            self.__model.AddNoOverlap(machine_to_intervals[machine])
        # Precedences inside a job.
        for job_id, job in enumerate(self.jobs_data):
            for task_id in range(len(job) - 1):
                # model.Add(all_tasks[job_id, task_id +1].start >= all_tasks[job_id, task_id].end)
                """限制條件: 各站運輸時間(輸送帶運送時間)"""
                self.__model.Add(all_tasks[job_id, task_id +1].start >= all_tasks[job_id, task_id].end + self.P2P_time[task_id])   
        for i in  range(len(self.jobs_data)):
            if (i != 0):
                self.__model.Add(all_tasks[i, 1].start >= all_tasks[0, 1].end + 46)
                self.__model.Add(all_tasks[i, 1].start >= all_tasks[i-1, 1].end + 46)
                self.__model.Add(all_tasks[i, 9].start >= all_tasks[i-1, 9].end + 46)
                self.__model.Add(all_tasks[i, 11].start >= all_tasks[i-1, 11].end + 46)
                self.__model.Add(all_tasks[0, 9].start >= all_tasks[i, 1].end + 46)
        # Makespan objective.
        obj_var = self.__model.NewIntVar(0, horizon, 'makespan')
        self.__model.AddMaxEquality(obj_var, [
            all_tasks[job_id, len(job) - 1].end
            for job_id, job in enumerate(self.jobs_data)
        ])
        self.__model.Minimize(obj_var)
        # Solve model.
        solver = cp_model.CpSolver()
        status = solver.Solve(self.__model)
        if status == cp_model.OPTIMAL:
            # Create one list of assigned tasks per machine.
            assigned_jobs = collections.defaultdict(list)
            for job_id, job in enumerate(self.jobs_data):
                for task_id, task in enumerate(job):
                    machine = task[0]
                    assigned_jobs[machine].append(
                        assigned_task_type(start=solver.Value(
                            all_tasks[job_id, task_id].start),
                                        job=job_id,
                                        index=task_id,
                                        duration=task[1]))
            # Create per machine output lines.
            output = ''
            for machine in all_machines:
                # Sort by starting time.
                assigned_jobs[machine].sort()
                if self.language == 'TW':
                    sol_line_tasks = self.__machine_name_tw[machine] + ' : ' # 中文版 FESTO Machine 顯示
                else :
                    sol_line_tasks = '%-10s : ' % self.__machine_name_en[machine] # 英文版 FESTO Machine 顯示
                sol_line = ' ' * 13
                for assigned_task in assigned_jobs[machine]:
                    name = 'job_%i_%i' % (assigned_task.job, assigned_task.index)
                    # Add spaces to output to align columns.
                    sol_line_tasks += '%-10s' % name
                    start = assigned_task.start
                    duration = assigned_task.duration
                    self.ScheduleData.append(task_data(machine, assigned_task.job +1, start, duration))
                    sol_tmp = '[%i,%i]' % (start, start + duration)
                    # Add spaces to output to align columns.
                    sol_line += '%-10s' % sol_tmp
                sol_line += '\n'
                sol_line_tasks += '\n'
                output += sol_line_tasks
                output += sol_line
            # Finally print the solution found.
            self.output = output
            if self.language == 'TW':
                print('預計總加工時間: %i 秒' % solver.ObjectiveValue())
            else :
                print('Optimal Schedule Length: %i Seconds' % solver.ObjectiveValue())
            # -----------Gatt plot-------------------------
            if self.GattPlot:
                final_data = {}  # 儲存排序後結果
                count = 0 # 記錄迴圈總共跑幾次
                where_order_at = ['A'] * len(self.jobs_data) # 記錄黑色小車在哪一區，用來判斷AGV去或回
                for machine in all_machines:
                    assigned_jobs[machine].sort()
                    for assigned_task in assigned_jobs[machine]:
                        start = assigned_task.start
                        duration = assigned_task.duration
                        final_data[count] = task_data(machine, assigned_task.job +1, start, duration) 
                        plt.barh(final_data[count].machine, final_data[count].duration, 
                                left = final_data[count].start_time,
                                color = self.__bar_color[final_data[count].oPos -1],
                                height = 0.9)
                        # ---判斷AGV去或回並顯示在甘特圖---
                        # Go: A->B，Back: B->A
                        if machine == 1:
                            if where_order_at[assigned_task.job] == 'A':
                                bar_text = "Go"
                                where_order_at[assigned_task.job] = 'B'
                            elif where_order_at[assigned_task.job] == 'B':
                                bar_text = "Back"
                                where_order_at[assigned_task.job] = 'A'

                            plt.text(final_data[count].start_time, 1, bar_text,color="white")
                        # --------------------------------
                        count += 1
                # 圖表多語系設定
                if self.language == 'TW':
                    for i in range(len(self.jobs_data)):
                        plt.barh(0,0,label = '工單 ' + str(final_data[i].oPos),color = self.__bar_color[final_data[i].oPos -1])
                    plt.xlabel('時間 (分鐘)')
                    plt.title('排程甘特圖 (總時間: ' + str(solver.ObjectiveValue()) + ' 秒 )')
                    plt.yticks(range(machine +1), self.__machine_name_tw)
                    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] #甘特圖中文字型設定
                    plt.rcParams['axes.unicode_minus'] = False
                else:
                    for i in range(len(self.jobs_data)):
                        plt.barh(0,0,label = 'Order ' + str(final_data[i].oPos),color = self.__bar_color[final_data[i].oPos -1])
                    plt.xlabel('Time (min.)')
                    plt.title('Gannt (Time Cost: ' + str(solver.ObjectiveValue()) + ' seconds )')
                    plt.yticks(range(machine +1), self.__machine_name_en)
                plt.xticks(range(0, int(solver.ObjectiveValue()) +60, 60), range(int(solver.ObjectiveValue()/60)+2))
                plt.xticks(range(0, int(solver.ObjectiveValue()) +60, 30))
                plt.grid(True, color = "grey", linewidth = "0.5", linestyle='dashed', axis = 'x') # 圖表格線設定
                plt.legend(loc = 'upper left')
                plt.show()            
    def WiteFile(self):
        path = os.getcwd()
        path +=  "\JobShopResult_" + str(order_num) + "單排程結果" + ".txt"  
        file = open(path , 'w')
        file.write(self.output)
        file.close
    def Layout_SchedulaData(self):
        JobShopDictionary = {"Storage":[],"AGV":[],"Check part":[],"Milling":[],"PCB":[],"Fuse check":[],"Cover":[],"Pressure":[],"Heat":[]}
        for var in self.ScheduleData:
            if var.machine == 0:
                JobShopDictionary["Storage"].append(var.oPos)
            if var.machine == 1:
                JobShopDictionary["AGV"].append(var.oPos)
            if var.machine == 2:
                JobShopDictionary["Check part"].append(var.oPos)
            if var.machine == 3:
                JobShopDictionary["Milling"].append(var.oPos)
            if var.machine == 4:
                JobShopDictionary["PCB"].append(var.oPos) 
            if var.machine == 5:
                JobShopDictionary["Fuse check"].append(var.oPos) 
            if var.machine == 6:
                JobShopDictionary["Cover"].append(var.oPos) 
            if var.machine == 7:
                JobShopDictionary["Pressure"].append(var.oPos) 
            if var.machine == 8:
                JobShopDictionary["Heat"].append(var.oPos) 
        print("輸出排程結果......\n")        
        print(JobShopDictionary)
        return JobShopDictionary

DL_cooling = DeeplearningModel.cooling_model()
DL_heating = DeeplearningModel.heating_model()
order_num = int(input("請輸入工單數目:  "))
aircon_temperture = int(input("請輸入冷氣溫度:  "))
 # task = (machine_id, processing_time).
jobs_data = [[(0, 18), (1, 53), (2, 2), (3, 7), (1, 61), (4, 72),(5, 1), (6, 2), (7, 3), (1, 61), (8, 76), (1,61), (0,15)],]
for i in range(order_num-1):
    jobs_data.append([(0, 18), (1, 61), (2, 2), (3, 7), (1, 61), (4, 72),(5, 1), (6, 2), (7, 3), (1, 61), (8, 76), (1,61), (0,15)])
heat_predict = DL_heating.model.predict([[25.5,30,aircon_temperture]])
jobs_data[0][10] = (8,int(heat_predict))
print("預測加熱機台加工時長")
print("預測次數: 1")
print("oPos= 1 加熱預測時間 =",int(heat_predict))
duration = 0
start1 = 0
start2 = 0
for count in range(order_num -1):
    model = MESJobShop()
    model.GattPlot = False
    model.jobs_data = jobs_data
    model.CPmodelSovel()
    print("預測次數:",count+2)
    for var in model.ScheduleData:
        if var.machine == 8 and var.oPos == count +1:
            duration = var.duration
            start1 = var.start_time
        if var.machine == 8 and var.oPos == count +2:
            start2 = var.start_time
            cool_predict = DL_cooling.model.predict([[30,start2 - start1 - duration ,aircon_temperture]])
            print('oPos= %i 加熱前溫度 = %i'%(count+2,cool_predict))
            temp = np.array([[cool_predict[0][0],30,aircon_temperture]],np.float)
            heat_predict = DL_heating.model.predict(temp)  
            print('oPos= %i 加熱預測時間 = %i'%(count+2,int(heat_predict)))
            jobs_data[count + 1][10] = (8,int(heat_predict))
    model.__del__()
model = MESJobShop()
model.GattPlot = True
model.jobs_data = jobs_data
model.CPmodelSovel()
print(model.output)
model.Layout_SchedulaData()
# model.WiteFile()
model.__del__()

