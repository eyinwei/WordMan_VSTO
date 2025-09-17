# MultiLevel 多级列表模块

## 文件结构

### 核心文件
- **MultiLevelListForm.cs** - 主多级列表窗体，包含完整的UI和业务逻辑
- **MultiLevelListForm.Designer.cs** - 主窗体的设计器文件

### 样式设置
- **LevelStyleSettingsForm.cs** - 多级段落样式设置窗体
- **LevelStyleSettingsForm.Designer.cs** - 样式设置窗体的设计器文件

### 数据管理
- **MultiLevelDataManager.cs** - 数据管理（合并了DataModels、ValidationHelper、ConfigurationManager）
- **WordStyleInfo.cs** - Word样式信息类

## 功能特性

### 1. 多级列表管理
- 支持1-9级列表设置
- 编号样式、格式、缩进设置
- 制表位和链接样式配置
- 实时预览和验证

### 2. 样式设置
- 字体、段落、缩进设置
- 颜色、对齐方式配置
- 段前分页等高级选项
- 实时样式预览

### 3. 配置管理
- CSV格式配置导入导出
- 配置验证和错误处理
- 默认配置生成

### 4. Word API集成
- 使用WordAPIHelper进行单位转换
- 通过Word API获取和设置样式
- 确保所有数值转换的准确性

## 使用说明

### 创建多级列表窗体
```csharp
var multiLevelListForm = new MultiLevelListForm();
multiLevelListForm.Show();
```

### 创建样式设置窗体
```csharp
var styleForm = new LevelStyleSettingsForm(maxLevel: 9);
styleForm.ShowDialog();
```

### 导入导出配置
```csharp
// 导出配置
ConfigurationManager.SaveConfigurationToFile(filePath, levelDataList, currentLevels);

// 导入配置
ConfigurationManager.LoadConfigurationFromFile(filePath, out levelDataList, out currentLevels);
```

## 代码质量

- 遵循Word API优先原则
- 使用统一的命名空间（WordMan_VSTO）
- 完整的错误处理和验证
- 代码简洁，避免冗余
- 良好的注释和文档

## 依赖关系

- **CustomControls.cs** - 提供StyledComboBox、StyledTextBox等自定义控件
- **WordAPIHelper.cs** - Word API统一管理
- **StylePane/ToggleButton.cs** - 切换按钮控件
