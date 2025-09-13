const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const ReportTemplate = sequelize.define('ReportTemplate', {
  id: {
    type: DataTypes.UUID,
    defaultValue: DataTypes.UUIDV4,
    primaryKey: true
  },
  name: {
    type: DataTypes.STRING,
    allowNull: false
  },
  originalFileName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  filePath: {
    type: DataTypes.STRING,
    allowNull: false
  },
  structure: {
    type: DataTypes.JSONB,
    allowNull: false
  },
  createdAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  },
  updatedAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  }
}, {
  tableName: 'report_templates',
  timestamps: true
});

const TemplateConfig = sequelize.define('TemplateConfig', {
  id: {
    type: DataTypes.UUID,
    defaultValue: DataTypes.UUIDV4,
    primaryKey: true
  },
  templateId: {
    type: DataTypes.UUID,
    references: {
      model: ReportTemplate,
      key: 'id'
    }
  },
  sectionId: {
    type: DataTypes.STRING,
    allowNull: false
  },
  sectionName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  dataType: {
    type: DataTypes.ENUM('fixed', 'manual', 'dynamic'),
    allowNull: false
  },
  value: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  sqlQuery: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  dataSourceId: {
    type: DataTypes.UUID,
    allowNull: true
  },
  createdAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  },
  updatedAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  }
}, {
  tableName: 'template_configs',
  timestamps: true
});

const DataSource = sequelize.define('DataSource', {
  id: {
    type: DataTypes.UUID,
    defaultValue: DataTypes.UUIDV4,
    primaryKey: true
  },
  name: {
    type: DataTypes.STRING,
    allowNull: false
  },
  type: {
    type: DataTypes.STRING,
    defaultValue: 'postgresql'
  },
  host: {
    type: DataTypes.STRING,
    allowNull: false
  },
  port: {
    type: DataTypes.INTEGER,
    allowNull: false
  },
  database: {
    type: DataTypes.STRING,
    allowNull: false
  },
  username: {
    type: DataTypes.STRING,
    allowNull: false
  },
  password: {
    type: DataTypes.STRING,
    allowNull: false
  },
  active: {
    type: DataTypes.BOOLEAN,
    defaultValue: true
  },
  createdAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  },
  updatedAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  }
}, {
  tableName: 'data_sources',
  timestamps: true
});

const ScheduledTask = sequelize.define('ScheduledTask', {
  id: {
    type: DataTypes.UUID,
    defaultValue: DataTypes.UUIDV4,
    primaryKey: true
  },
  templateId: {
    type: DataTypes.UUID,
    references: {
      model: ReportTemplate,
      key: 'id'
    }
  },
  name: {
    type: DataTypes.STRING,
    allowNull: false
  },
  cronExpression: {
    type: DataTypes.STRING,
    allowNull: false
  },
  active: {
    type: DataTypes.BOOLEAN,
    defaultValue: true
  },
  lastRunTime: {
    type: DataTypes.DATE,
    allowNull: true
  },
  nextRunTime: {
    type: DataTypes.DATE,
    allowNull: true
  },
  createdAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  },
  updatedAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  }
}, {
  tableName: 'scheduled_tasks',
  timestamps: true
});

const GeneratedReport = sequelize.define('GeneratedReport', {
  id: {
    type: DataTypes.UUID,
    defaultValue: DataTypes.UUIDV4,
    primaryKey: true
  },
  templateId: {
    type: DataTypes.UUID,
    references: {
      model: ReportTemplate,
      key: 'id'
    }
  },
  taskId: {
    type: DataTypes.UUID,
    references: {
      model: ScheduledTask,
      key: 'id'
    },
    allowNull: true
  },
  reportName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  filePath: {
    type: DataTypes.STRING,
    allowNull: false
  },
  generatedAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  },
  status: {
    type: DataTypes.ENUM('success', 'failed', 'pending'),
    defaultValue: 'pending'
  }
}, {
  tableName: 'generated_reports',
  timestamps: false
});

// Define associations
ReportTemplate.hasMany(TemplateConfig, { foreignKey: 'templateId', as: 'configs' });
TemplateConfig.belongsTo(ReportTemplate, { foreignKey: 'templateId' });

ReportTemplate.hasMany(ScheduledTask, { foreignKey: 'templateId', as: 'tasks' });
ScheduledTask.belongsTo(ReportTemplate, { foreignKey: 'templateId' });

ReportTemplate.hasMany(GeneratedReport, { foreignKey: 'templateId', as: 'reports' });
GeneratedReport.belongsTo(ReportTemplate, { foreignKey: 'templateId' });

ScheduledTask.hasMany(GeneratedReport, { foreignKey: 'taskId', as: 'reports' });
GeneratedReport.belongsTo(ScheduledTask, { foreignKey: 'taskId' });

TemplateConfig.belongsTo(DataSource, { foreignKey: 'dataSourceId' });
DataSource.hasMany(TemplateConfig, { foreignKey: 'dataSourceId' });

module.exports = {
  sequelize,
  ReportTemplate,
  TemplateConfig,
  DataSource,
  ScheduledTask,
  GeneratedReport
};