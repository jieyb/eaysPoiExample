package com.example.service.impl;

import com.example.entity.MyTest;
import com.example.service.MyTestService;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class MyTestServiceImpl implements MyTestService {
    @Override
    public List<MyTest> getDataList() {
        List<MyTest> list = new ArrayList<>();
        MyTest myTest = new MyTest();
        myTest.setName("小红");
        myTest.setSex(2);
        myTest.setAge(22);
        myTest.setMobile("15778192883");
        list.add(myTest);

        MyTest myTest2 = new MyTest();
        myTest2.setName("小黑");
        myTest2.setSex(1);
        myTest2.setAge(22);
        myTest2.setMobile("15778192882");
        list.add(myTest2);

        return list;
    }
}
