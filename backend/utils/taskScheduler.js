const cron = require('node-cron');

class TaskScheduler {
  constructor(prisma, reportGenerator) {
    this.prisma = prisma;
    this.reportGenerator = reportGenerator;
    this.scheduledJobs = new Map();
  }
  
  async loadActiveTasks() {
    try {
      const tasks = await this.prisma.rs_scheduled_tasks.findMany({
        where: { active: true }
      });
      
      for (const task of tasks) {
        this.scheduleTask(task);
      }
      
      console.log(`Loaded ${tasks.length} active scheduled tasks`);
    } catch (error) {
      console.error('Error loading scheduled tasks:', error);
    }
  }
  
  scheduleTask(task) {
    // Cancel existing job if any
    this.cancelTask(task.id);
    
    // Validate cron expression
    if (!cron.validate(task.cronExpression)) {
      console.error(`Invalid cron expression for task ${task.id}: ${task.cronExpression}`);
      return;
    }
    
    // Schedule new job
    const job = cron.schedule(task.cronExpression, async () => {
      console.log(`Running scheduled task: ${task.name}`);
      
      try {
        // Update last run time
        await this.prisma.rs_scheduled_tasks.update({
          where: { id: task.id },
          data: {
            lastRunTime: new Date(),
            nextRunTime: this.getNextRunTime(task.cronExpression)
          }
        });
        
        // Generate report
        await this.reportGenerator.generateReport(task.templateId, task.id);
        
        console.log(`Scheduled task completed: ${task.name}`);
      } catch (error) {
        console.error(`Error running scheduled task ${task.id}:`, error);
      }
    });
    
    this.scheduledJobs.set(task.id, job);
    
    // Update next run time
    this.prisma.rs_scheduled_tasks.update({
      where: { id: task.id },
      data: {
        nextRunTime: this.getNextRunTime(task.cronExpression)
      }
    }).catch(console.error);
  }
  
  cancelTask(taskId) {
    const job = this.scheduledJobs.get(taskId);
    if (job) {
      job.stop();
      this.scheduledJobs.delete(taskId);
    }
  }
  
  getNextRunTime(cronExpression) {
    try {
      const interval = cron.parseExpression(cronExpression);
      return interval.next().toDate();
    } catch (error) {
      console.error('Error parsing cron expression:', error);
      return null;
    }
  }
}

module.exports = TaskScheduler;