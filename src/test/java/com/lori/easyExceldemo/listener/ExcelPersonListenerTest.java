package com.lori.easyExceldemo.listener;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.lori.easyExceldemo.model.Person;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Map;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelPersonListenerTest {

    @Test
    public void testExcelPersonListener() {

        try(InputStream inputStream = new FileInputStream("src/main/resources/excel/person.xlsx")) {
            AnalysisEventListener listener = new ExcelPersonListener();

            ExcelReader reader = new ExcelReader(inputStream,null,listener);

            /**
             * @see Sheet 第一个参数是第几个sheet解析，第二个参数是从第几行开始读取
             *
             * reader.read执行时同步的,需要等解析完成
             *
             * 业务场景：合法数据保存，不合法数据返回异常信息，可以通过listener返回回来。
             * 比如：只能保存一千条，其他的返回错误信息
             */
            Long startTime = System.currentTimeMillis();
            reader.read(new Sheet(1,2,Person.class));
            Long endTime = System.currentTimeMillis();

            System.out.println("解析总时间 ： "+(endTime - startTime));

            Assert.assertTrue(30 == ((ExcelPersonListener) listener).getFailCount());

            Assert.assertTrue(1000 == ((ExcelPersonListener) listener).getSuccessCount());

            for (Map.Entry<Integer,String> entry : ((ExcelPersonListener) listener).getMap().entrySet()){
                System.out.println("错误的行数: "+entry.getKey()+"==="+"错误的原因: "+entry.getValue());
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
