using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaskScheduler;

namespace UserElevation
{
    class TaskSchedulerClient
    {
        TaskScheduler.TaskScheduler objScheduler;
        //To hold Task Definition
        ITaskDefinition objTaskDef;
        //To hold Trigger Information
        IRegistrationTrigger objTrigger;
        //To hold Action Information
        IExecAction objAction;
        const string taskName = "ElevatedTask";

        public void CreatTask()
        {
            try
            {
                objScheduler = new TaskScheduler.TaskScheduler();
                objScheduler.Connect();

                //Setting Task Definition
                SetTaskDefinition();
                //Setting Task Trigger Information
                //SetTriggerInfo();
                //Setting Task Action Information
                SetActionInfo();

                //Getting the roort folder
                ITaskFolder root = objScheduler.GetFolder("\\");
                //Registering the task, if the task is already exist then it will be updated
                IRegisteredTask regTask = root.RegisterTaskDefinition(taskName, objTaskDef, (int)_TASK_CREATION.TASK_CREATE_OR_UPDATE, null, null, _TASK_LOGON_TYPE.TASK_LOGON_INTERACTIVE_TOKEN, "");

                //To execute the task immediately calling Run()
                IRunningTask runtask = regTask.Run(null);

                Console.WriteLine("Task is created successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        //Setting Task Definition
        private void SetTaskDefinition()
        {
            try
            {
                objTaskDef = objScheduler.NewTask(0);

                objTaskDef.Principal.RunLevel = _TASK_RUNLEVEL.TASK_RUNLEVEL_HIGHEST;
                
                //Registration Info for task
                //Name of the task Author
                objTaskDef.RegistrationInfo.Author = "A cat";
                //Description of the task 
                objTaskDef.RegistrationInfo.Description = taskName;
                //Registration date of the task 
                objTaskDef.RegistrationInfo.Date = DateTime.Today.ToString("yyyy-MM-ddTHH:mm:ss"); //Date format 

                //Settings for task
                //Thread Priority
                objTaskDef.Settings.Priority = 7;
                //Enabling the task
                objTaskDef.Settings.Enabled = true;
                //To hide/show the task
                objTaskDef.Settings.Hidden = false;
                //Execution Time Lmit for task
                //objTaskDef.Settings.ExecutionTimeLimit = "PT10M"; //10 minutes
                //Specifying no need of network connection
                objTaskDef.Settings.RunOnlyIfNetworkAvailable = false;


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Setting Task Trigger Information
        private void SetTriggerInfo()
        {
            try
            {
                //Trigger information based on time - TASK_TRIGGER_TIME
                objTrigger = (IRegistrationTrigger)objTaskDef.Triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_REGISTRATION);
                //Trigger ID
                objTrigger.Id = taskName + "Trigger";

                //Start Time

                //objTrigger.StartBoundary = "2014-01-09T10:10:00"; //yyyy-MM-ddTHH:mm:ss
                //End Time

                //objTrigger.EndBoundary = "2016-01-01T07:30:00"; //yyyy-MM-ddTHH:mm:ss
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Setting Task Action Information
        private void SetActionInfo()
        {
            try
            {
                //Action information based on exe- TASK_ACTION_EXEC
                objAction = (IExecAction)objTaskDef.Actions.Create(_TASK_ACTION_TYPE.TASK_ACTION_EXEC);
                //Action ID
                objAction.Id = taskName;
                //Set the path of the exe file to execute, Here mspaint will be opened
                objAction.Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CheckAndRepairTask(ref IRegisteredTask task)
        {
            objTaskDef = task.Definition;
            var actionsEnum = objTaskDef.Actions.GetEnumerator();

            while (actionsEnum.MoveNext())
            {
                IExecAction action = actionsEnum.Current as IExecAction;

                if (!File.Exists(action.Path))
                {

                    objTaskDef.Actions.Remove(actionsEnum.Current);
                    SetActionInfo();
                }

                break;
            }
        }

        public void DeleteTask()
        {
            try
            {
                TaskScheduler.TaskScheduler objScheduler = new TaskScheduler.TaskScheduler();
                objScheduler.Connect();

                ITaskFolder containingFolder = objScheduler.GetFolder("\\");
                //Deleting the task
                containingFolder.DeleteTask(taskName, 0);  //Give name of the Task
                
                Console.WriteLine("Task deleted...");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public bool TryGetTask(out IRegisteredTask task)
        {
            task = null;
            TaskScheduler.TaskScheduler objScheduler = new TaskScheduler.TaskScheduler();
            objScheduler.Connect();

            ITaskFolder containingFolder = objScheduler.GetFolder("\\");
            var tasks = containingFolder.GetTasks(0);
            var tasksEnum = tasks.GetEnumerator();
            IRegisteredTask regTask;

            while (tasksEnum.MoveNext() && task == null)
            {
                regTask = tasksEnum.Current as IRegisteredTask;

                if (regTask.Name == taskName)
                    task = regTask;
            }

            return task != null;
        }
    }
}
