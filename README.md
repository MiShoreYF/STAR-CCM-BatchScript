STAR-CCM+ 批量案例生成脚本

作用：
1. 读取Excel参数配置表，基于Java宏模板和SIM文件模板自动生成多个仿真案例
2. 通过config.ini配置文件配置参数映射和替换规则
3. 可添加多个需要替换参数的模板

使用流程：
1. 准备CasePlan.xlsx算例规划参数表
2. 准备template_Case.sim算例模板文件
3. 通过STARCCM+录制template_Macro.java宏模板文件
4. 准备其他需要替换参数的模板文件（可选）
3. 修改config.ini配置文件，配置参数映射和替换规则
4. 运行脚本

注意事项：
1. 需要安装第三方依赖库：pandas
2. Excel文件需包含标题行，数据从第二行开始
3. 模板文件与此脚本在同一目录下
4. 所有要进行批量替换的模板文件都要以 template_ 开头，并放在脚本所在目录
5. 需要手动修改 template_Macro.java 中待替换的字符，包括：
   - 模板文件名必须修改为 template_Macro
   - 计算结果自动导出的文件名必须修改为 CaseName
   - 算例保存路径必须修改为 SavePath
   - 待替换的值必须修改为不与宏文件中其他字符冲突的占位符

**STAR-CCM+ Batch Case Generation Script**  

**Purpose**:  
1. Read an Excel parameter configuration table and automatically generate multiple simulation cases based on Java macro templates and SIM file templates.  
2. Configure parameter mappings and replacement rules via the `config.ini` configuration file.  
3. Support adding multiple template files requiring parameter replacements.  

**Workflow**:  
1. Prepare the `CasePlan.xlsx` case planning parameter table.  
2. Prepare the `template_Case.sim` case template file.  
3. Record the `template_Macro.java` macro template file via STAR-CCM+.  
4. Prepare other template files requiring parameter replacements (optional).  
5. Modify the `config.ini` file to configure parameter mappings and replacement rules.  
6. Run the script.          

**Notes**:  
1. Requires installation of third-party dependency libraries: `pandas`.  
2. The Excel file must include a header row, with data starting from the second row.  
3. Template files must be located in the same directory as the script.  
4. All template files requiring batch replacements must start with `template_` and be placed in the script directory. 
5. Manual modifications are required in `template_Macro.java`:  
   - The template file name must be set to `template_Macro`.  
   - The auto-exported result file name must be set to `CaseName`.  
   - The case save path must be set to `SavePath`.  
   - Replaceable values must use placeholders that do not conflict with other characters in the macro file.  
