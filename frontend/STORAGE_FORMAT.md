# 表格配置存储格式文档

## 数据库表结构
```sql
rs_template_configs {
  id: String (UUID)
  templateId: String
  sectionId: String  -- 配置项的唯一标识
  sectionName: String -- 配置项的名称描述
  dataType: FIXED | DYNAMIC | MANUAL
  value: String -- 存储的值
  sqlQuery: String -- SQL查询语句
  dataSourceId: String -- 数据源ID外键
  columnIndex: Number -- 列索引
  parentSectionId: String -- 父级section ID
}
```

## sectionId 格式规范

### 1. 表格级别配置
- **表格名称**: `${sectionId}_name`
  - 例：`table_section_0_name`
  - value: 表格标题

- **数据源**: `${sectionId}_datasource` 
  - 例：`table_section_0_datasource`
  - value: 数据源ID，dataSourceId: 同样的ID

- **备注**: `${sectionId}_remark`
  - 例：`table_section_0_remark` 
  - value: 备注文本

- **统计开关**: `${sectionId}_stats`
  - 例：`table_section_0_stats`
  - value: 'true' | 'false'

### 2. 单元格配置
- **普通单元格**: `cell_${sectionId}_${rowIndex}_${colIndex}`
  - 例：`cell_table_section_0_1_2`
  - value: 根据displayType存储不同格式

- **统计行单元格**: `stats_${sectionId}_${colIndex}`
  - 例：`stats_table_section_0_1`
  - value: 根据fillType存储

### 3. 合并单元格配置
- **合并单元格**: `merge_${sectionId}_${mergeKey}`
  - 例：`merge_table_section_0_1_2`
  - mergeKey: `${startRow}_${startCol}` 
  - value: JSON字符串存储合并信息
  - 格式：`{"startRow":1,"startCol":2,"rowspan":2,"colspan":3}`

## 单元格配置 value 格式

### 文字模式 (displayType: 'text')
- dataType: 'FIXED'
- value: 直接存储文字内容
- 例：`value: "666"`

### 数据集模式 (displayType: 'dataset')
- dataType: 'DYNAMIC'
- value: `"${dataStructure}|${displayFields.join(',')}"`
- dataSourceId: 数据集ID
- 例：`value: "single|name,age"`, `dataSourceId: "1"`

## 恢复逻辑
1. 按 sectionId 前缀/后缀过滤配置
2. 将每种配置类型恢复到对应的 reactive 对象
3. 确保配置间不会互相覆盖

## 关键点
- sectionId 必须唯一且不重叠
- 单元格配置使用 `cell_` 前缀确保与其他配置分离
- value 格式要统一，便于解析
- 恢复时先清空目标对象，避免污染