from pathlib import Path
from copy import deepcopy
import zipfile, hashlib, json, re, math
from lxml import etree as E
from PIL import Image, ImageDraw, ImageFont

BASE = Path(__file__).resolve().parent.parent
WORK = BASE / '.work'
SOURCE = WORK / 'original_202310232005.docx'
FINAL = WORK / 'basic_candidate.docx'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W, 'r': R, 'm': M}
q = lambda x: '{'+W+'}'+x

# All twenty entries were checked against publisher pages, society pages,
# author manuscripts, or university repositories on 2026-11-10.
refs = [
('杨世春, 崔海港, 周思达, 等. 车控操作系统实时性/安全性保障技术[J]. 北京航空航天大学学报, 2024, 50(7): 2051-2065.', '10.13700/j.bh.1001-5965.2022.0594'),
('李易庭, 高泽鹏, 李梦梦, 等. 实现纹波电流还原的PMSM端口模拟算法[J]. 汽车工程, 2024, 46(8): 1469-1478.', '10.19562/j.chinasae.qcgc.2024.08.013'),
('章回炫, 范涛, 边元均, 等. 永磁同步电机高性能电流预测控制[J]. 电工技术学报, 2022, 37(17): 4335-4345.', '10.19595/j.cnki.1000-6753.tces.211054'),
('连传强, 郭力源, 许观达, 等. 永磁伺服系统转速跟踪控制与数据驱动的参数设计方法[J]. 电工技术学报, 2024, 39(8): 2412-2421.', '10.19595/j.cnki.1000-6753.tces.230163'),
('田里思, 宗耐龙, 刘立伟, 等. 基于单电流传感器的永磁同步电机混合脉宽调制技术[J]. 电工技术学报, 2025, 40(8): 2451-2463.', '10.19595/j.cnki.1000-6753.tces.240519'),
('马铱林, 袁浩, 尹威, 等. 考虑等效电磁损耗电阻偏移的永磁同步电机直流信号注入在线参数辨识方法[J]. 电工技术学报, 2023, 38(22): 6015-6026.', '10.19595/j.cnki.1000-6753.tces.230176'),
('李婕, 杨淑英, 谢震, 等. 基于有效信息迭代快速粒子群优化算法的永磁同步电机参数在线辨识[J]. 电工技术学报, 2022, 37(18): 4604-4613.', '10.19595/j.cnki.1000-6753.tces.211470'),
('杨帆, 赵希梅, 金鸿雁, 等. 基于广义比例积分观测器的永磁同步电机鲁棒谐振预测电流控制[J]. 电工技术学报, 2024, 39(10): 2995-3006.', '10.19595/j.cnki.1000-6753.tces.230331'),
('东野亚兰, 杨淑英, 王奇帅, 等. 基于增强型扩张状态观测器的永磁同步电机低抖振高抗扰二阶终端滑模电流控制[J]. 电工技术学报, 2024, 39(8): 2434-2448.', '10.19595/j.cnki.1000-6753.tces.222359'),
('陈荣, 翟凯淼, 舒胡平. 永磁同步电机双矢量固定开关频率模型预测控制研究[J]. 电工技术学报, 2023, 38(14): 3812-3823.', '10.19595/j.cnki.1000-6753.tces.221857'),
('汪逸哲, 黄晟, 廖武, 等. 基于新型虚拟矢量调制方法的IPMSM模型预测电流控制方法[J]. 电工技术学报, 2024, 39(8): 2422-2433.', '10.19595/j.cnki.1000-6753.tces.230105'),
('吴翔, 陈硕, 李佳, 等. 基于改进正交锁相环的永磁同步电机无位置传感器控制[J]. 电工技术学报, 2024, 39(2): 475-486.', '10.19595/j.cnki.1000-6753.tces.221926'),
('赵文祥, 宋世昌, 周书文, 等. 改进滑模观测器的电流源逆变器驱动PMSM无位置传感器控制[J]. 电工技术学报, 2024, 39(4): 987-995.', '10.19595/j.cnki.1000-6753.tces.222146'),
('魏尧, 柯栋梁, 黄东晓, 等. 基于超局部化时间序列的永磁同步电机无模型预测电流滑模控制策略[J]. 电工技术学报, 2024, 39(4): 1022-1032.', '10.19595/j.cnki.1000-6753.tces.230301'),
('王祥松, 杨淑英, 谢震, 等. 基于多参数辨识的永磁同步电机磁链预测控制[J]. 电工技术学报, 2024, 39(22): 7084-7098.', '10.19595/j.cnki.1000-6753.tces.231680'),
('HU M, AHN H, PARK J, et al. Novel rapid control prototyping for permanent magnet synchronous motor via model-based design and STM32 chip[J]. The International Journal of Advanced Manufacturing Technology, 2024, 135: 1187-1204.', '10.1007/s00170-024-14579-4'),
('SU Z, WANG D, YU Z, et al. PHCG: Optimizing Simulink code generation for embedded system with SIMD instructions[J]. IEEE Transactions on Computer-Aided Design of Integrated Circuits and Systems, 2023, 42(4): 1072-1084.', '10.1109/TCAD.2022.3197504'),
('CHATRI C, LABBADI M, OUASSAID M, et al. Design and implementation of finite-time control for speed tracking of permanent magnet synchronous motors[J]. IEEE Control Systems Letters, 2023, 7: 721-726.', '10.1109/LCSYS.2022.3221719'),
('BOSSO A, TILLI A, CONFICONI C. A hybrid sensorless observer for the robust global asymptotic flux reconstruction of permanent magnet synchronous machines[J]. IEEE Control Systems Letters, 2022, 6: 3367-3372.', '10.1109/LCSYS.2022.3184649'),
('ZHANG Z, LIU X, YU J, et al. Time-varying disturbance observer based improved sliding mode single-loop control of PMSM drives with a hybrid reaching law[J]. IEEE Transactions on Energy Conversion, 2023, 38(4): 2539-2549.', '10.1109/TEC.2023.3277628'),
]
(WORK/'references.json').write_text(json.dumps([{'number':i+1,'language':'中文' if i<15 else '英文','citation':a,'doi':b,'url':'https://doi.org/'+b} for i,(a,b) in enumerate(refs)],ensure_ascii=False,indent=2),encoding='utf-8')

background = [
('lead','选题背景及意义、国内外研究现状、研究方法和思路等，不少于2000字，并列出参考文献。'),
('h','1 选题背景及意义'),
('p','永磁同步电机（Permanent Magnet Synchronous Motor，PMSM）具有较高的功率密度和运行效率，在新能源汽车电驱动及机电执行机构中具有重要应用价值。电机控制器需要根据转矩或转速指令，将直流电源转换为幅值、频率和相位受控的三相电压，同时完成电流检测、转子位置获取和异常保护。磁场定向控制（Field-Oriented Control，FOC）通过坐标变换，将定子电流分解为励磁分量和转矩分量，为实现平稳调速和快速电流调节提供了清晰的控制结构。'),
('p','对于本科阶段的电机控制设计，仅在仿真中得到理想波形，尚不能说明控制算法能够在实际硬件上稳定运行。采样偏置、逆变器死区、位置角零偏、离散延时及处理器执行时间，都可能造成模型与样机表现不一致。如果算法模型、嵌入式代码和测试用例分别维护，后续调参和问题定位也容易失去依据。因此，本课题将控制性能与开发过程的一致性同时作为研究对象，在STM32平台上建立能够复现、比较和验证的PMSM矢量控制系统。'),
('p','本课题借鉴车企电控开发中按需求分解、分层实现和逐级验证的思路，采用基于模型的设计（Model-Based Design，MBD）方法，将需求、控制模型、生成代码和试验结果关联起来。杨世春等对汽车控制软件实时性和验证技术的讨论，以及李易庭等对电机端口模拟精度的研究，说明实时调度与被控对象模型的可信度均应进入验证范围[1-2]。本课题据此构建适用于低压电机实验的开发流程，使每项性能指标都有对应模型、测试工况和验收记录。'),
('p','课题拟构建低压小功率最小FOC系统，完成电源、STM32最小系统、逆变器、采样保护和上位机通信设计，优先使用编码器并兼容霍尔接口。通过模型在环、软件在环、处理器执行验证与电机台架试验，研究车企MBD方法在原型开发中的应用，为后续无位置传感器控制提供基础。'),
]

research = [
('h','2 国内外研究现状'),
('sub','2.1 国内研究现状'),
('p','在电流调节和速度跟踪方面，研究重点已从理想控制器设计扩展到数字实现约束。章回炫等针对永磁同步电机电流预测控制研究时序优化，强调计算延时对动态性能的影响[3]。连传强等通过滞后—超前补偿和数据驱动的参数设计改善永磁伺服系统的低频转速跟踪性能[4]。田里思等研究单电流传感器下的混合脉宽调制及电流重构，表明采样时刻和可观测区处理直接影响电流反馈质量[5]。这些工作启示本课题：应先保证采样和PWM更新时序正确，再评价PI参数的优劣；为降低初次实现的复杂度，拟采用多电阻采样方案，暂不将单电阻重构作为必做内容。'),
('p','在参数变化和非理想因素处理方面，马铱林等将损耗电阻变化纳入在线参数辨识模型[6]；李婕等通过有效信息迭代改进粒子群算法，以减少参数辨识的计算量[7]。杨帆等使用广义比例积分观测器和谐振预测控制处理参数失配与死区扰动[8]，东野亚兰等将增强型扩张状态观测器与二阶终端滑模电流控制结合，兼顾抖振抑制和抗扰性能[9]。这些成果说明，控制模型需要保留参数变化和逆变器非线性的影响。本课题拟优先完成电阻、电感、极对数及磁链参数核对，并在仿真中引入参数摄动，检验基线FOC方案的适用范围。'),
('p','在预测控制方面，陈荣等通过缩小电压矢量选择范围和优化作用时间计算，实现固定开关频率的双矢量控制[10]；汪逸哲等采用新的虚拟矢量调制方法，在改善电流质量的同时降低搜索计算负担[11]。预测控制具有处理多变量目标的潜力，但其模型、离散实现和执行开销也需要配套设计。因此，本课题将其作为与PI型FOC对照的研究背景，重点完成结构清晰、便于生成代码和定位问题的速度—电流双闭环，而不同时叠加多种高级算法。'),
('p','在无位置传感器控制方面，吴翔等研究改进正交锁相环的转子信息获取方法[12]；赵文祥等针对电流源逆变器驱动中的谐波影响，改进滑模观测器和滤波环节[13]。相关结论体现了位置估计对信号质量和驱动拓扑的依赖，不能直接跨拓扑照搬。魏尧等利用超局部化时间序列及在线估计构建无模型预测电流滑模控制[14]，王祥松等将多参数辨识和磁链预测结合，提高参数变化条件下的控制适应性[15]。这类方法为后续提升鲁棒性提供了方向，也进一步凸显建立可靠基线平台和统一对比工况的必要性。'),
('sub','2.2 国外研究现状'),
('p','在MBD与快速原型方面，Hu等将MATLAB/Simulink、STM32CubeMX与FOC相结合，在STM32平台验证了有传感器和无传感器控制模型，为从模型到嵌入式样机的实现提供了直接参考[16]。Su等研究Simulink模型的代码生成优化，说明自动生成代码仍需关注目标架构和执行效率[17]。前者侧重电机实验开发链路，后者侧重生成代码的计算开销；二者共同支持本课题将功能正确性与实时性分开检查，并对模型接口、数据类型和任务周期进行统一约束。'),
('p','在高性能调速与观测方面，Chatri等研究考虑负载扰动的有限时间转速跟踪控制，并进行实时实验[18]；Bosso等利用混合系统方法研究PMSM磁链重构的稳定性，给出了相应激励条件[19]；Zhang等通过混合趋近律和时变扰动观测器研究速度—电流单环控制[20]。这些工作分别从动态响应、状态估计和抗扰结构出发提升性能，但实验平台、控制周期和传感条件不同，其性能结论需要结合具体工况理解。本课题采用双闭环PI作为基准，并记录扰动响应和资源占用，为后续算法比较保留统一条件。'),
('sub','2.3 研究评述与课题定位'),
('p','综上，近年PMSM研究已形成电流调节、参数辨识、位置观测及快速原型等多个方向。对于本课题，关键在于将这些研究中反复出现的采样质量、模型误差、计算延时和验证条件落实到可完成的毕业设计任务中。拟以编码器反馈、基速以下运行和PI型FOC建立基线，以需求追踪和分层测试组织开发；先进控制、无感启动和弱磁运行作为拓展方向，不作为基本验收的前置条件。'),
('h','3 研究方法和思路'),
('sub','3.1 借鉴车企MBD流程组织开发'),
('p','采用V形开发与验证思路：左侧由系统需求分解为硬件、控制算法和软件接口需求，中间通过离散模型与自动代码生成完成实现，右侧按单元、软件集成、控制器和电机系统逐级验证。为需求建立唯一编号，保存需求条目、模型模块、代码函数、测试用例及结果的对应关系。每阶段以可审查的产物作为进入下一阶段的依据；测试发现偏差时，回溯到参数、模型或驱动接口并执行回归测试。'),
('fig','mbd'),
('cap','图1 适用于本课题的MBD开发与验证流程'),
('p','模型在环（MIL）阶段，在同一仿真环境中连接控制器与电机、逆变器模型，验证算法和工况覆盖；软件在环（SIL）阶段，将生成的C代码在主机上运行，比较其与离散模型的输出；处理器在环（PIL）阶段，将算法代码放到STM32执行，由主机传递测试向量并核对结果，同时测量目标端执行时间。硬件在环（HIL）需要实时运行的被控对象模拟器及确定性的I/O闭环，普通串口监测和接入真实电机不构成HIL。本课题以MIL、SIL、PIL及实际电机台架为实施主线，HIL仅在具备设备时补充开展。'),
('sub','3.2 PMSM建模与双闭环FOC设计'),
('p','在三相绕组对称、反电动势近似正弦的条件下，建立同步旋转坐标系中的电压、转矩和机械运动模型。采用幅值不变的Clarke、Park变换，定义d轴与转子永磁磁链对齐，电角速度与机械角速度满足ωe=pωm。模型的基本关系为：'),
('eq','u_d = R_s i_d + L_d di_d/dt − ω_e L_q i_q'),
('eq','u_q = R_s i_q + L_q di_q/dt + ω_e (L_d i_d + ψ_f)'),
('eq','T_e = (3/2)p[ψ_f i_q + (L_d − L_q)i_d i_q]'),
('eq','J dω_m/dt = T_e − T_L − Bω_m'),
('p','式中，Rs为相电阻，Ld、Lq为d、q轴电感，ψf为永磁磁链，p为极对数，J为转动惯量，B为黏性阻尼系数，TL为负载转矩。电机参数从铭牌、资料及低压测量获得，并记录测量方法和单位。以表贴式PMSM为优先对象，在额定电流和电压约束内采用id*=0控制；若实际电机存在显著凸极性，模型保留Ld与Lq的差异，最大转矩电流比控制另行扩展。'),
('p','速度误差经速度PI得到iq*，并通过限流器约束；采样得到的相电流经Clarke和Park变换形成id、iq反馈，两个电流PI输出调节电压。加入交叉耦合与反电动势前馈后，经电压矢量限幅、逆Park变换及空间矢量脉宽调制（SVPWM）生成三相占空比。SVPWM直接由αβ电压计算开关作用时间，无需在其后再串接独立的逆Clarke变换。速度给定采用斜坡，两个环节均设置抗积分饱和；逆变器死区和调制限幅在仿真与驱动中保持一致。'),
('fig','foc'),
('cap','图2 编码器反馈的速度与电流双闭环FOC结构'),
('p','先整定电流环，再闭合速度环。忽略耦合并进行前馈补偿时，以Rs和Ld、Lq估计初始PI参数，按Kp,d=Ldωci、Kp,q=Lqωci、Ki=Rsωci设置，再依据离散延时和台架响应修正。电流环初始带宽取500 Hz量级，速度环取20 Hz量级，最终通过仿真与试验确认。PWM和电流控制频率拟为20 kHz，速度环为1 kHz，电流控制周期为50 μs；通过速率转换和明确的信号更新次序避免跨周期数据混用。'),
('sub','3.3 硬件与嵌入式软件实现'),
('p','硬件以12 V直流母线和STM32G431为初步方案。电源采用12 V降压至5 V、再由LDO获得3.3 V，分别为驱动及数字模拟电路供电；STM32最小系统包含去耦、晶振、复位、启动配置、SWD和状态指示。功率级由三相栅极驱动器与六只MOSFET构成，结合器件导通损耗、耐压裕量和散热条件选型，设置互补PWM、死区、驱动欠压及硬件关断通路。母线电容和能量吸收措施依据减速时的回馈能量核算，避免只按稳态输入功率选型。'),
('p','反馈电路以U、V两相电流采集为基础，预留第三相低侧采样通道，按照有效导通窗口选择可用电流并重构剩余相电流；独立母线电流用于监视和过流保护，同时采集母线电压及功率级温度。采样电阻、放大倍数和偏置按最大电流及ADC输入范围计算，采用开尔文取样、适当滤波与功率回路和采样回路分区布线。编码器A/B/Z与霍尔接口兼容预留，主方案采用编码器定位；USB转串口用于MATLAB上位机调参、状态监视和波形记录。'),
('p','软件划分为硬件驱动层、信号适配层、控制算法层和应用管理层。控制算法层使用固定步长、单精度和清晰的输入输出接口，通过Embedded Coder生成C代码；ADC标定、编码器计数、PWM装载和串口收发由驱动层实现。PWM定时器触发ADC，转换完成后执行快速控制任务，在约定更新点装载占空比。状态机设置待机、自检、校准、运行、停机和故障等状态，故障锁存后需满足恢复条件并重新使能；通信任务通过缓冲区和降采样运行，不阻塞电流中断。'),
('sub','3.4 对比验证与结果分析'),
('p','建立启动、给定转速变化、负载阶跃、母线波动、参数摄动和反馈异常用例。MIL中分别加入ADC偏置与噪声、PWM死区、一个控制周期量级的延时及电机参数变化，比较理想模型与含非理想因素模型的差异；SIL和PIL使用相同输入、初值和采样序列，比较限幅前后控制电压、状态切换和占空比。台架按电源检查、PWM检查、采样标定、转子定向、电流闭环、速度闭环的顺序逐级上电，并保留原始数据、测试条件和版本信息。'),
('p','性能评价采用稳态转速误差、阶跃超调量、调节时间、负载扰动恢复时间、电流跟踪均方根误差和控制任务执行时间。速度相对误差以给定值为分母，零速工况单独报告绝对误差；调节时间按进入并保持在给定转速±2%带内计算。电流THD作为补充观察量，说明采样带宽、稳态窗口和开关纹波处理方法，避免混用统计口径。通过统一工况比较，分析PI整定、限幅与抗饱和、采样时序和参数变化对系统表现的影响。'),
('h','参考文献'),
]

tasks = [
('lead','毕业论文（设计）的主要内容（列出具体任务）、预期目标及拟解决的关键问题。'),
('h','1 主要内容'),
('p','（1）需求分析与方案设计。梳理低压PMSM矢量控制的功能、接口、运行边界和性能指标，建立需求追踪表及测试计划；确定基速以下、编码器反馈的速度—电流双闭环方案，形成系统框图、信号清单和软硬件任务划分。'),
('p','（2）控制建模与仿真。建立PMSM、三相逆变器、采样链路和负载模型，实现Clarke/Park变换、双电流PI、速度PI、前馈解耦、限幅抗饱和和SVPWM。使用离散固定步长模型完成启动、调速、负载阶跃及参数摄动测试，形成参数表、工况脚本和MIL结果。'),
('p','（3）硬件设计与调试。完成12 V/5 V/3.3 V电源、STM32最小系统、三相驱动逆变、电流电压采样、编码器及霍尔接口、USB转串口等原理图与PCB设计，输出器件清单；核算功率器件、采样范围及保护阈值，开展分模块检查后再接电机调试。'),
('p','（4）代码生成与集成。统一模型信号单位、数据类型、初值和执行顺序，生成控制算法C代码；完成驱动适配、状态管理及上位机数据协议，通过SIL和PIL比较模型与代码输出，测量快速控制路径执行时间，保存代码生成配置和可重建的工程版本。'),
('p','（5）系统试验与论文整理。在一致的电机、电源和负载条件下进行速度跟踪及抗扰试验，比较理想仿真、含非理想因素仿真与实测波形，分析误差来源。形成模型、硬件资料、软件工程、测试记录和毕业论文，并将未覆盖的工况与后续改进方向说明清楚。'),
('h','2 预期目标'),
('p','以下指标为设计和验收目标，需通过仿真与台架测试验证。拟选12 V、额定功率不高于50 W且满足目标转速范围的PMSM，驱动相电流限制初步不高于5 A；具体连续电流、热限制及负载能力按最终电机和功率板额定值确定。若选定器件边界不足，应在详细设计阶段调整指标并同步更新需求，不能通过超额定运行满足测试。'),
('metric','R01 功能与接口：实现上电自检、静态采样校准、编码器零位标定、受控启动、正向调速、斜坡停机及故障锁存；具备电流、电压、转速和故障状态监测能力。反向运行与再生制动在供电及能量吸收条件确认后开展。'),
('metric','R02 实时性：PWM和电流环20 kHz，速度环1 kHz；快速路径最坏实测执行时间不超过35 μs，并确认ADC转换、计算和PWM更新均落入对应时序窗口。记录Flash/RAM占用、最大中断延时及是否发生漏周期。'),
('metric','R03 稳态调速：在母线12 V、设定转速300～1500 r/min、负载不超过额定转矩50%的测试范围内，稳态平均转速相对误差目标不大于1%。进入稳态后取不少于1 s窗口，每个工况重复3次，报告均值和最大偏差。'),
('metric','R04 动态响应：在母线12 V、无外加负载下，进行500→1000 r/min给定阶跃，速度超调量目标不大于10%，按±2%误差带计算的调节时间目标不大于0.5 s。试验全过程保持电流限幅，超出限流能力时注明该工况不满足指标前提。'),
('metric','R05 抗扰与电流调节：在1000 r/min时施加0→30%额定转矩的负载阶跃，速度偏差峰值目标不大于给定值10%，0.5 s内恢复至±2%误差带；在给定电流非零且未饱和工况，d、q轴电流跟踪误差RMSE分别不大于额定相电流峰值的5%。'),
('metric','R06 一致性与保护：SIL/PIL在相同测试向量下，归一化占空比与模型的最大绝对偏差目标不大于10⁻⁴，离散状态切换序列一致。过流由比较器或驱动故障信号直接触发PWM硬件关断，目标响应不超过一个PWM周期；欠压、过压、超温和编码器超时通过信号注入验证，阈值按器件额定值与测量误差设定。'),
('p','验证上述指标时，转速以编码器测量并用独立转速测量手段复核；负载试验使用可测量转矩的制动器或对拖装置。若缺少可计量的加载条件，R05在MIL中完成定量验证，台架只报告实际可测的负载响应及其局限，不以手动摩擦或无标定负载替代规定转矩阶跃。'),
('h','3 拟解决的关键问题'),
('p','（1）模型与硬件接口的一致性。统一电流符号、相序、电角度方向、编码器计数与极对数换算；采用静态定向和低电流旋转验证角度零偏。建立物理量、标幺量及ADC原始值之间的换算关系，避免量纲和比例错误造成反馈正反馈或电流失控。'),
('p','（2）采样与PWM协同。针对低侧采样在部分占空比下有效窗口不足的问题，设计与中心对齐PWM同步的ADC触发，避开开关瞬态，并核算最小采样窗口。必要时限制调制深度或使用预留第三相通道；对通道零偏、增益误差和电流重构一致性分别标定。'),
('p','（3）双闭环稳定性与电压饱和。针对限流、母线波动和电压矢量饱和，设计分级限幅、抗积分饱和及给定斜坡。通过含延时模型检查相位裕度与响应，先固定速度调电流环，再闭合速度环，防止两个环节带宽接近而引发振荡。'),
('p','（4）自动代码的可部署性。明确单精度、固定步长、多速率交互和中断执行约束；将硬件相关代码封装在驱动层，避免外设访问混入通用控制模型。以SIL/PIL测试检查数值误差，以GPIO翻转和周期计数测量执行开销，确保通信、监视及后台任务不破坏快速控制周期。'),
('p','（5）验证结果的可追踪性。每项测试保存需求编号、模型与代码版本、配置参数、输入工况和原始波形；对偏差先区分控制算法、模型参数、采样链路和时序问题，再进行修正与回归。预期形成的特色是可复现的MBD工程闭环及其STM32实现，不预先宣称提出新的控制理论。'),
]

outline = [
('lead','毕业论文（设计）提纲'),
('olh','1 绪论'),('ol','1.1 研究背景与意义'),('ol','1.2 国内外研究现状'),('ol','1.3 研究内容与技术路线'),
('olh','2 系统需求与MBD开发方案'),('ol','2.1 功能需求与性能指标'),('ol','2.2 系统架构与软硬件接口'),('ol','2.3 V形开发流程与需求追踪'),
('olh','3 PMSM矢量控制建模与仿真'),('ol','3.1 PMSM数学模型与坐标变换'),('ol','3.2 双闭环PI及前馈解耦设计'),('ol','3.3 SVPWM与离散时序'),('ol','3.4 MIL工况及非理想因素分析'),
('olh','4 STM32控制系统硬件设计'),('ol','4.1 电源与最小系统'),('ol','4.2 三相逆变器与驱动保护'),('ol','4.3 电流电压采样与位置接口'),('ol','4.4 PCB设计与模块调试'),
('olh','5 代码生成与嵌入式软件实现'),('ol','5.1 算法模型接口与代码生成配置'),('ol','5.2 驱动适配与状态机管理'),('ol','5.3 SIL与PIL一致性验证'),('ol','5.4 上位机通信与实时性分析'),
('olh','6 台架试验与结果分析'),('ol','6.1 平台与测试方法'),('ol','6.2 稳态调速与动态响应'),('ol','6.3 负载扰动与保护验证'),('ol','6.4 仿真与实测偏差分析'),
('olh','7 总结与展望'),('ol','7.1 工作总结'),('ol','7.2 不足与后续改进'),
]

plans = [
('2026.11.10-2026.11.30','检索文献，明确MBD流程、控制范围及指标，完成开题报告与需求清单。'),
('2026.12.1-2026.12.15','核对电机参数与器件方案，建立PMSM及逆变器模型，确定接口和测试计划。'),
('2026.12.16-2027.1.1','完成双闭环FOC、限幅抗饱和及SVPWM模型，开展MIL及参数摄动测试。'),
('2027.1.2-2027.1.20','完成硬件原理图、PCB及物料核算，检查采样窗口、驱动与保护方案。'),
('2027.1.21-2027.2.10','完成控制代码生成与SIL对比，集成驱动和状态机，检查硬件各模块。'),
('2027.2.11-2027.2.28','开展PIL及执行时间测量，完成电流标定、编码器定向和低压电流闭环。'),
('2027.3.1-2027.4.4','完成速度闭环、上位机与台架测试，分析偏差并回归验证，整理试验数据。'),
('2027.4.5-2027.4.21','完成论文初稿、图表和设计资料归档，按指导意见修改与完善。'),
('2027.4.22-2027.5.12','完成论文定稿、演示与答辩材料，复核系统和资料，准备毕业答辩。'),
]

exec((WORK/'basic_scope.py').read_text(encoding='utf-8'), globals())

def diagrams():
    font=lambda n:ImageFont.truetype('C:/Windows/Fonts/simsun.ttc',n)
    def canvas(h):
        im=Image.new('RGBA',(1800,h),(255,255,255,0));return im,ImageDraw.Draw(im)
    def text(d,xy,s,size=34):
        lines=s.split('\n');f=font(size)
        for k,line in enumerate(lines):d.text((xy[0],xy[1]+k*(size+10)),line,font=f,fill='black',anchor='mm')
    def box(d,xy,s,size=34):
        d.rectangle(xy,outline='black',width=3);n=len(s.split('\n'));text(d,((xy[0]+xy[2])/2,(xy[1]+xy[3])/2-(n-1)*(size+10)/2),s,size)
    def arr(d,pts,dash=False):
        if not dash:d.line(pts,fill='black',width=3)
        else:
            for a,b in zip(pts,pts[1:]):
                dist=math.hypot(b[0]-a[0],b[1]-a[1]);steps=int(dist/16)
                for j in range(0,steps,2):
                    t1=j/max(steps,1);t2=min(j+1,steps)/max(steps,1);d.line([(a[0]+(b[0]-a[0])*t1,a[1]+(b[1]-a[1])*t1),(a[0]+(b[0]-a[0])*t2,a[1]+(b[1]-a[1])*t2)],fill='black',width=2)
        a,b=pts[-2:];ang=math.atan2(b[1]-a[1],b[0]-a[0]);d.polygon([b,(b[0]-17*math.cos(ang-.45),b[1]-17*math.sin(ang-.45)),(b[0]-17*math.cos(ang+.45),b[1]-17*math.sin(ang+.45))],fill='black')
    im,d=canvas(760)
    left=[(70,55,580,155,'系统方案\n速度环与电流环'),(190,260,700,360,'控制模型\nPI与坐标变换'),(310,465,820,565,'仿真调试\n启动与调速')]
    right=[(1220,55,1730,155,'电机实验\n转速与电流波形'),(1100,260,1610,360,'硬件调试\n采样与PWM输出'),(980,465,1490,565,'代码生成\nSTM32程序集成')]
    for x1,y1,x2,y2,s in left+right:box(d,(x1,y1,x2,y2),s)
    arr(d,[(325,155),(445,260)]);arr(d,[(445,360),(565,465)]);arr(d,[(820,515),(980,515)])
    arr(d,[(1235,465),(1355,360)]);arr(d,[(1355,260),(1475,155)])
    arr(d,[(580,105),(1220,105)],True);arr(d,[(700,310),(1100,310)],True)
    text(d,(900,650),'模型仿真、代码实现与电机实验相互校核',38)
    text(d,(900,710),'根据运行波形调整PI参数与采样设置',32)
    im.save(WORK/'mbd.png')
    im,d=canvas(690)
    xs=[(60,140,300,255,'速度PI\n电流指令限幅'),(390,140,660,255,'d/q电流PI\n电压指令限幅'),(750,140,990,255,'逆Park\n变换'),(1080,140,1320,255,'SVPWM\n占空比'),(1410,140,1730,255,'逆变器\nPMSM')]
    for a,b,c,e,s in xs:box(d,(a,b,c,e),s,32)
    for i in range(4):arr(d,[(xs[i][2],197),(xs[i+1][0],197)])
    text(d,(175,65),'转速给定 n*',34);arr(d,[(175,90),(175,140)])
    text(d,(525,70),'id*=0    iq*来自速度PI',30)
    box(d,(985,410,1320,510),'ADC与Clarke/Park\n相电流反馈',32)
    arr(d,[(1570,255),(1570,460),(1320,460)]);arr(d,[(985,460),(525,460),(525,255)])
    box(d,(1430,565,1730,655),'编码器\n角度与速度',32)
    arr(d,[(1670,255),(1770,255),(1770,610),(1730,610)])
    arr(d,[(1430,610),(175,610),(175,255)])
    text(d,(740,565),'速度反馈 n',30)
    arr(d,[(1430,610),(1370,610),(1370,355),(870,355),(870,255)])
    arr(d,[(1150,355),(1150,410)])
    text(d,(1000,320),'电角度 θe',30)
    im.save(WORK/'foc.png')

diagrams()
with zipfile.ZipFile(SOURCE) as z:
    original={n:z.read(n) for n in z.namelist()}
    infos=z.infolist()
root=E.fromstring(original['word/document.xml'])
rels=E.fromstring(original['word/_rels/document.xml.rels'])
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
def relation(target,typ,external=False):
    rid='rIdReport'+str(len(rels)+1)
    a=E.SubElement(rels,'{'+PKG+'}Relationship',Id=rid,Type=R+'/'+typ,Target=target)
    if external:a.set('TargetMode','External')
    return rid

tables=root.find(q('body')).findall(q('tbl'))
rows=tables[1].findall(q('tr'))
template_h=rows[1].find(q('tc')).findall(q('p'))[1]
template_p=rows[1].find(q('tc')).findall(q('p'))[2]

def el(tag,**attrs):
    return E.Element(q(tag),{q(k):str(v) for k,v in attrs.items()})

def paragraph(s,kind='p'):
    p=E.Element(q('p'))
    proto=template_h if kind=='h' else template_p
    pr=deepcopy(proto.find(q('pPr')))
    if pr is None:pr=E.Element(q('pPr'))
    p.append(pr)
    for name in ['numPr','pageBreakBefore','sectPr','tabs']:
        for n in pr.findall(q(name)):pr.remove(n)
    # Explicit line pitch and source fonts, preserving the form's 12 pt body.
    for name in ['ind','spacing','jc','keepNext','keepLines','snapToGrid']:
        for n in pr.findall(q(name)):pr.remove(n)
    pr.append(el('snapToGrid',val='0'))
    if kind not in ['h']:
        for n in pr.findall(q('pStyle')):pr.remove(n)
    size=24;line=360;before=0;after=0
    align='both'
    if kind=='h':before=156;after=100;pr.append(el('keepNext'));pr.append(el('keepLines'))
    elif kind=='sub':before=100;after=40;pr.append(el('keepNext'));pr.append(el('keepLines'))
    elif kind in ['lead','olh']:align='left';pr.append(el('keepNext'));pr.append(el('keepLines'));after=80
    elif kind=='ol':align='left';line=280
    elif kind=='ref':size=21;line=260;after=0;pr.append(el('keepLines'))
    elif kind=='cap':size=21;line=300;align='center';after=80
    elif kind=='eq':size=23;line=330;align='center'
    elif kind=='metric':line=340;after=40
    pr.append(el('spacing',before=before,after=after,line=line,lineRule='auto'))
    pr.append(el('jc',val=align))
    if kind in ['p','metric']:pr.append(el('ind',firstLine='480'))
    elif kind=='ol':pr.append(el('ind',left='240'))
    elif kind=='ref':pr.append(el('ind',left='360',hanging='360'))
    else:pr.append(el('ind',firstLine='0'))
    if kind=='olh':line=280;after=10
    # Direct false values override any inherited pagination in the source.
    if kind not in ['h','sub','olh','lead']:
        pr.append(el('widowControl',val='1'))
        pr.append(el('keepNext',val='1' if kind=='eq' else '0'))
        pr.append(el('keepLines',val='0' if kind in ['p','metric'] else '1'))
    def run(text):
        r=E.SubElement(p,q('r'));rp=E.SubElement(r,q('rPr'))
        rp.append(el('rFonts',ascii='Times New Roman',hAnsi='Times New Roman',eastAsia='宋体',cs='Times New Roman'))
        rp.append(el('color',val='000000'));rp.append(el('sz',val=size));rp.append(el('szCs',val=size))
        if kind in ['h','sub','olh']:rp.append(el('b'))
        t=E.SubElement(r,q('t'));t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=text
        return r
    if kind=='eq':
        # Native Word OMML, using actual mathematical subscript structures.
        omp=E.SubElement(p,'{'+M+'}oMathPara');om=E.SubElement(omp,'{'+M+'}oMath')
        def mr(parent,t):
            rr=E.SubElement(parent,'{'+M+'}r');E.SubElement(rr,'{'+M+'}t').text=t
        for token in re.split(r'([A-Za-zωψ]_[A-Za-z])',s):
            if re.fullmatch(r'[A-Za-zωψ]_[A-Za-z]',token):
                a,b=token.split('_');ss=E.SubElement(om,'{'+M+'}sSub');mr(E.SubElement(ss,'{'+M+'}e'),a);mr(E.SubElement(ss,'{'+M+'}sub'),b)
            elif token:mr(om,token)
    else:run(s)
    return p

def imagep(name):
    rid=relation('media/report_'+name+'.png','image')
    original['word/media/report_'+name+'.png']=(WORK/(name+'.png')).read_bytes()
    p=paragraph('','cap');p.find(q('pPr')).append(el('keepNext'))
    wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';a='http://schemas.openxmlformats.org/drawingml/2006/main';pic='http://schemas.openxmlformats.org/drawingml/2006/picture'
    im=Image.open(WORK/(name+'.png'));cx=5050000;cy=int(cx*im.height/im.width)
    r=E.SubElement(p,q('r'));dr=E.SubElement(r,q('drawing'));inline=E.SubElement(dr,'{'+wp+'}inline',distT='0',distB='0',distL='0',distR='0')
    E.SubElement(inline,'{'+wp+'}extent',cx=str(cx),cy=str(cy));E.SubElement(inline,'{'+wp+'}docPr',id='301' if name=='mbd' else '302',name='Report '+name,descr='MBD开发流程' if name=='mbd' else '速度电流双闭环FOC结构')
    graphic=E.SubElement(inline,'{'+a+'}graphic');gd=E.SubElement(graphic,'{'+a+'}graphicData',uri=pic);pi=E.SubElement(gd,'{'+pic+'}pic')
    nv=E.SubElement(pi,'{'+pic+'}nvPicPr');E.SubElement(nv,'{'+pic+'}cNvPr',id='0',name=name+'.png');E.SubElement(nv,'{'+pic+'}cNvPicPr')
    fill=E.SubElement(pi,'{'+pic+'}blipFill');E.SubElement(fill,'{'+a+'}blip',{'{'+R+'}embed':rid});st=E.SubElement(fill,'{'+a+'}stretch');E.SubElement(st,'{'+a+'}fillRect')
    sp=E.SubElement(pi,'{'+pic+'}spPr');xf=E.SubElement(sp,'{'+a+'}xfrm');E.SubElement(xf,'{'+a+'}off',x='0',y='0');E.SubElement(xf,'{'+a+'}ext',cx=str(cx),cy=str(cy));geom=E.SubElement(sp,'{'+a+'}prstGeom',prst='rect');E.SubElement(geom,'{'+a+'}avLst')
    return p

def fill_cell(tc,items):
    for n in list(tc):
        if n.tag!=q('tcPr'):tc.remove(n)
    idx=0
    while idx<len(items):
        kind,s=items[idx]
        if kind=='eq':
            # One native equation array keeps the four basic motor equations
            # together when a table continues onto the next page.
            expressions=[]
            while idx<len(items) and items[idx][0]=='eq':
                expressions.append(items[idx][1]);idx+=1
            p=paragraph('', 'eq');om=p.find('.//{'+M+'}oMath')
            for child in list(om):om.remove(child)
            array=E.SubElement(om,'{'+M+'}eqArr')
            for expression in expressions:
                entry=E.SubElement(array,'{'+M+'}e')
                source_math=paragraph(expression,'eq').find('.//{'+M+'}oMath')
                for child in source_math:entry.append(deepcopy(child))
            pr=p.find(q('pPr'))
            for n in pr.findall(q('keepNext')):pr.remove(n)
            pr.append(el('keepNext',val='0'))
            tc.append(p)
            continue
        tc.append(imagep(s) if kind=='fig' else paragraph(s,kind))
        idx+=1
    pr=tc.find(q('tcPr'))
    va=pr.find(q('vAlign'))
    if va is not None:va.set(q('val'),'top')
    # Large narrative slots must be allowed to flow across pages.
    tr=tc.getparent();trpr=tr.find(q('trPr'))
    if trpr is not None:
        for n in trpr.findall(q('cantSplit')):trpr.remove(n)
        trpr.append(el('cantSplit',val='0'))

fill_cell(rows[1].find(q('tc')),background)
# A keep-with-next chain in a single merged cell can make Word move the
# complete introductory row. Its two headings are already at the row start.
for p in rows[1].findall('.//'+q('p')):
    pr=p.find(q('pPr'))
    for n in pr.findall(q('keepNext')):pr.remove(n)
    pr.append(el('keepNext',val='0'))
fill_cell(rows[2].find(q('tc')),research)
ref_cell=rows[2].find(q('tc'))
for i,(citation,doi) in enumerate(refs,1):
    p=paragraph(f'[{i}] {citation} DOI: ', 'ref')
    h=E.SubElement(p,q('hyperlink'),{'{'+R+'}id':relation('https://doi.org/'+doi,'hyperlink',True)})
    r=E.SubElement(h,q('r'));rp=E.SubElement(r,q('rPr'));rp.append(el('rFonts',ascii='Times New Roman',hAnsi='Times New Roman',eastAsia='宋体'));rp.append(el('sz',val='21'));rp.append(el('color',val='000000'));E.SubElement(r,q('t')).text=doi+'.'
    ref_cell.append(p)
fill_cell(rows[3].find(q('tc')),tasks)
fill_cell(rows[4].find(q('tc')),outline)
# The schedule is a distinct page pattern in both supplied references.
for p in rows[5].findall('.//'+q('p')):
    pr=p.find(q('pPr'))
    if pr is None:pr=E.SubElement(p,q('pPr'))
    for name in ['pageBreakBefore','keepNext']:
        for n in pr.findall(q(name)):pr.remove(n)
    pr.append(el('keepNext'))
for p in rows[6].findall('.//'+q('p')):
    pr=p.find(q('pPr'))
    if pr is None:pr=E.SubElement(p,q('pPr'))
    for n in pr.findall(q('keepNext')):pr.remove(n)
    pr.append(el('keepNext'))
for i,(dates,task) in enumerate(plans,7):
    cs=rows[i].findall(q('tc'))
    fill_cell(cs[1],[('schedule',dates)])
    fill_cell(cs[2],[('schedule',task)])
    for tc in cs:
        for p in tc.findall(q('p')):
            pr=p.find(q('pPr'))
            if pr is None:pr=E.SubElement(p,q('pPr'))
            for n in pr.findall(q('spacing')):pr.remove(n)
            pr.append(el('spacing',before='30',after='30',line='280',lineRule='auto'));pr.append(el('snapToGrid',val='0'))

# The teacher's choice is not a student-supplied fact; remove the inherited tick.
for row in rows[29:31]:
    for t in row.iter(q('t')):
        if t.text:t.text=t.text.replace('☑','□')

# Separate the introductory paragraphs into source-derived continuation rows.
# This avoids Word reserving a whole nearly-page-sized merged cell after
# the project-title row. The borders remain one continuous form section.
intro=rows[1]
intro_cell=intro.find(q('tc'))
intro_paras=list(intro_cell.findall(q('p')))
intro_position=list(tables[1]).index(intro)
tables[1].remove(intro)
for idx,p in enumerate(intro_paras):
    tr=deepcopy(intro)
    tc=tr.find(q('tc'))
    for child in list(tc):
        if child.tag!=q('tcPr'):tc.remove(child)
    tc.append(deepcopy(p))
    borders=tc.find(q('tcPr')).find(q('tcBorders'))
    for side,visible in [('top',idx==0),('bottom',idx==len(intro_paras)-1)]:
        b=borders.find(q(side))
        if b is None:b=E.SubElement(borders,q(side))
        b.set(q('val'),'single' if visible else 'nil')
    for h in tr.findall('./'+q('trPr')+'/'+q('trHeight')):h.getparent().remove(h)
    tables[1].insert(intro_position+idx,tr)

# Keep the review heading with its following form instead of orphaning it
# below the advisor signature page, as in the unfilled source.
body=root.find(q('body'))
review_seen=False
for child in body:
    if child.tag==q('p') and '开题报告评审纪要' in ''.join(child.itertext()):
        review_seen=True
        pr=child.find(q('pPr'))
        if pr is None:pr=E.SubElement(child,q('pPr'))
        pr.append(el('pageBreakBefore'))
        pr.append(el('keepNext'))
        for pos in child.findall('.//'+q('position')):pos.getparent().remove(pos)
    elif review_seen and child.tag==q('p'):
        pr=child.find(q('pPr'))
        if pr is None:pr=E.SubElement(child,q('pPr'))
        pr.append(el('keepNext'))
    elif review_seen and child.tag==q('tbl'):break

# Start the schedule as a fresh source-derived table so that its title and
# column headings cannot land on separate pages. Two unused spacer rows are
# removed to retain the one-page schedule pattern after adding its title.
schedule_table=E.Element(q('tbl'))
for name in ['tblPr','tblGrid']:
    schedule_table.append(deepcopy(tables[1].find(q(name))))
for row in rows[5:]:
    tables[1].remove(row)
    if row not in [rows[16],rows[17]]:schedule_table.append(row)
sep=E.Element(q('p'));seppr=E.SubElement(sep,q('pPr'))
seppr.append(el('pageBreakBefore'));seppr.append(el('spacing',before='0',after='0',line='20',lineRule='exact'))
seppr.append(el('keepNext'))
place=list(body).index(tables[1])+1
body.insert(place,sep);body.insert(place+1,schedule_table)

# Leave all advisor opinions, signatures, and reviewer information untouched.
# All other package parts remain byte-for-byte identical to the retained source.
# Canonical OOXML ordering is necessary for Word's table pagination flags.
porder='pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split()
torder='cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden ins del trPrChange'.split()
for parent in root.iter(q('pPr')):
    parent[:]=sorted(parent,key=lambda n:porder.index(E.QName(n).localname) if E.QName(n).localname in porder else 100)
for parent in root.iter(q('trPr')):
    parent[:]=sorted(parent,key=lambda n:torder.index(E.QName(n).localname) if E.QName(n).localname in torder else 100)
original['word/document.xml']=E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
original['word/_rels/document.xml.rels']=E.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
with zipfile.ZipFile(FINAL,'w',zipfile.ZIP_DEFLATED) as z:
    old_names=set()
    for inf in infos:z.writestr(inf,original[inf.filename]);old_names.add(inf.filename)
    for n,data in original.items():
        if n not in old_names:z.writestr(n,data)

content=''.join(s for _,s in background+research+tasks+outline)
(WORK/'body.txt').write_text(content,encoding='utf-8')
print(json.dumps({'output':str(FINAL),'body_chinese_characters':len(re.findall(r'[\u4e00-\u9fff]',content)),'references':len(refs),'chinese_refs':15,'english_refs':5,'source_sha256':hashlib.sha256(SOURCE.read_bytes()).hexdigest()},ensure_ascii=False))
