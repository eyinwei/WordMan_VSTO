# 样式设置窗格 (StylePane)

这是一个复刻自参考代码的Word样式设置窗格，提供了完整的样式管理功能。

## 功能特性

### 1. 样式管理
- 添加新样式
- 删除现有样式
- 修改样式属性
- 保存/加载样式配置

### 2. 样式属性设置
- **字体设置**：字体名称、大小、颜色、加粗、斜体、下划线
- **段落设置**：对齐方式、行距、段前距、段后距、首行缩进
- **编号设置**：编号样式、编号格式、起始编号

### 3. 内置样式支持
- 支持Word内置样式（标题1-9、正文、目录等）
- 自动获取系统可用字体
- 智能单位转换（字符、磅、厘米）

## 文件结构

```
StylePane/
├── StyleSettingsForm.cs          # 主窗体逻辑
├── StyleSettingsForm.Designer.cs # 窗体设计器代码
├── WordStyleInfo.cs              # 样式信息数据模型
├── StyleSerializationHelper.cs   # 样式序列化工具
├── NumericUpDownWithUnit.cs      # 带单位数值输入控件
├── ToggleButton.cs               # 切换按钮控件
└── README.md                     # 说明文档
```

## 使用方法

1. 在Word中点击"样式设置"按钮打开窗格
2. 选择要修改的样式或添加新样式
3. 调整字体、段落、编号等属性
4. 点击"应用样式"将设置应用到文档
5. 使用"保存配置"保存当前设置

## 技术特点

- 完全基于Word API实现，确保兼容性
- 支持XML序列化，配置可持久化
- 自定义控件提供更好的用户体验
- 遵循VSTO开发最佳实践

## 依赖项

- Microsoft.Office.Interop.Word
- System.Windows.Forms
- System.Drawing
- System.Xml.Serialization

## 注意事项

- 需要Word 2016或更高版本
- 某些功能需要Word文档处于活动状态
- 样式修改会立即应用到当前文档
