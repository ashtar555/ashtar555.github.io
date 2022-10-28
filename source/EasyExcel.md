# 应用场景：

使用java语言对Excel文件读/写处理

## Excel文件内容

![image-20220911220709870](https://xingqiu-tuchuang-1256524210.cos.ap-shanghai.myqcloud.com/7005/image-20220911220709870.png)

## 需求

1. 上传excel文件，将文件内容映射成学生对象，传递给前端展示或者存入数据库
2. 下载excel表格的学生数据信息

# 项目概览以及提前准备

## **注意点：**

>  必须使用Java8！！！

## 项目依赖

```xml
    <dependencies>
        <!-- Springboot依赖 -->
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-web</artifactId>
        </dependency>
        <!-- 热部署依赖 -->
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-devtools</artifactId>
            <scope>runtime</scope>
            <optional>true</optional>
        </dependency>
        <!-- lombok 作用:@Data 一键生成get和set方法 -->
        <dependency>
            <groupId>org.projectlombok</groupId>
            <artifactId>lombok</artifactId>
            <optional>true</optional>
        </dependency>
        <!-- Springboot整合junit -->
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-test</artifactId>
            <scope>test</scope>
            <exclusions>
                <exclusion>
                    <groupId>org.junit.vintage</groupId>
                    <artifactId>junit-vintage-engine</artifactId>
                </exclusion>
            </exclusions>
        </dependency>
        <!-- EasyExcel依赖 -->
        <dependency>
            <groupId>com.alibaba</groupId>
            <artifactId>easyexcel</artifactId>
            <version>2.2.10</version>
        </dependency>
    </dependencies>
```

## domain学生类

**@ExcelProperty(value = #, index = #)**  

value值与Excel表格列标题对应，如果有多层列标题，采用数组的形式。(若只有value属性可省略value)

index为列的次序(从0开始)

**@DateTimeFormat()**

定义Date类型的格式

@DateTimeFormat("yyyy-MM-dd") 将表格数据转换为例如:2022-09-11

```java
@Data
public class Student {
    @ExcelProperty(value = {"学生信息表","学生姓名"})
    private String stuId;
    @ExcelProperty({"学生信息表","学生姓名"})
    private String stuName;
    @ExcelProperty({"学生信息表","学生性别"})
    private String stuSex;
    @ExcelProperty({"学生信息表","学生出生日期"})
    @DateTimeFormat("yyyy-MM-dd")
    private Date birthday;
}
```

## 通用响应数据返回类Result

```java
@Data
public class Result<T> implements Serializable {

    private Integer code;
    private T data;
    private String msg;

    public Result() {

    }

    public Result(Integer code, T data) {
        this.code = code;
        this.data = data;
    }

    public Result(Integer code, T data, String msg) {
        this.code = code;
        this.data = data;
        this.msg = msg;
    }
}
```

## mapper 和 service层

简化项目，此处并未用到

思路：得到List<Student> 遍历得到每一个对象Student后，insert加入数据库或者直接批量插入数据库（可自己拓展）

## controller层

简单的创建StudentController

访问上传文件路径为:<a>localhost:8080/student/upload</a>

访问下载文件路径为:<a>localhost:8080/student/download</a>

## util工具

简单的创建ExcelUtil，其函数是简单打印下数据（可拓展功能）

```java
public class ExcelUtil {
    public static void readExcel(List<Student> studentList) {
        //对得到的数据进行简单处理
        for (Student student : studentList) {
            System.out.println("Stu = " + student.toString());
        }
    }
}
```

# Excel以及EasyExcel理解

一个Excel文件就是一个工作簿workBook，一个工作簿有许多个工作表Sheet

EasyExcel相当于一行一行通过**数据监听器**读取数据，省去了内存

## 数据监听器

```java
public class DataListener extends AnalysisEventListener<Student> {

    ArrayList<Student> students = new ArrayList<>();

    //数据一行一行读取时执行该函数
    @Override
    public void invoke(Student student, AnalysisContext analysisContext) {
        students.add(student);
        
        //设置每五个学生信息一起取出来
        //缺点：若不能%5.会有些数据读不到
        if (students.size() % 5 == 0) {
            //对取出来的这五条学生信息数据处理
            //ExcelUtil.readExcel() 该函数是自己编写的，功能是简单的打印学生信息
            ExcelUtil.readExcel(students);
            students.clear();
        }
    }

    //数据全部读取完后执行该函数
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("读取完毕");
    }
}
```

# 上传文件（读文件）

```java
@RestController
@RequestMapping("/student")
public class StudentController {

    @PostMapping("/upload")
    public Result<Boolean> readExcel(MultipartFile uploadExcel) {
        try {
            //EasyExcel为入口
            //EasyExcel.read()调用读函数
            //参数为：1.读取文件的文件名(或者inputStream)，2.提取出来的类型，3.数据监听器
            ExcelReaderBuilder readerBuilder = EasyExcel.read(uploadExcel.getInputStream(), Student.class, new DataListener());
            //readerBuilder.sheet()返回指定工作表sheet
            //参数为：1.sheet的序号(推荐)，2.sheet的名字(不推荐)，3.不写默认第一个
            ExcelReaderSheetBuilder sheet = readerBuilder.sheet();
            //得到sheet调用doRead函数
            sheet.doRead();
            //返回成功数据类型响应
            return new Result<>(1, true, "success");
        } catch (IOException e) {
            throw new RuntimeException(e);
            //...
            //返回失败数据类型响应
        }
    }
}
```

# 下载文件（写文件）

先创建一些数据出来，后续方便验证是否成功写入excel

```java
// mock一些数据 用于模拟
private List<Student> initData() {
    List<Student> students = new ArrayList<>();
    for (int i = 1; i <= 10; i++) {
        Student student = new Student();
        student.setStuId("123" + i);
        student.setStuName("yyf" + i);
        student.setStuSex("女");
        student.setBirthday(new Date());
        students.add(student);
    }
    return students;
}
```

和读操作一样**三步走**，但值得注意的是：写操作需要设置返回的请求头信息(直接复制即可)

```java
    @GetMapping("/download")
    public Result<Boolean> writeExcel(HttpServletResponse response) throws IOException {
        //mock些数据
        List<Student> students = initData();

        //固定设置返回头，复制粘贴即可
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        String fileName = URLEncoder.encode("学生信息表", "UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + fileName + ".xlsx");

		//EasyExcel入口 
        //EasyExcel.write()调用写函数，参数为：1.写的文件的文件名(或者outputStream)，2.提取出来的类型，3.数据监听器
        ExcelWriterBuilder writeBuilder = EasyExcel.write(response.getOutputStream(), Student.class);
        //写在工作表中，若无参数则为默认值
        ExcelWriterSheetBuilder sheet = writeBuilder.sheet();
        //调用写函数传数据
        sheet.doWrite(students);

        return new Result<>(1, true, "success");
    }
```

# gitee仓库

包含源代码，excel测试文档，该笔记markdown文档

https://gitee.com/Ashtar555/use-of-easy-excel

# 学习文档：

EasyExcel官方地址：https://easyexcel.opensource.alibaba.com/

推荐B站黑马视频：https://www.bilibili.com/video/BV1C7411275q?p=2&spm_id_from=333.880.my_history.page.click