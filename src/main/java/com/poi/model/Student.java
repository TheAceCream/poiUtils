package com.poi.model;

import java.util.Date;

/**
 * Created with IntelliJ IDEA.
 * Description:
 * User: weicaijia
 * Date: 2018/8/13 11:25
 * Time: 14:15
 */
public class Student {

    private Long id;

    private String name;

    private Byte sex;

    private Date birth;

    private Double mathGrade;

    private Boolean visible;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Byte getSex() {
        return sex;
    }

    public void setSex(Byte sex) {
        this.sex = sex;
    }

    public Date getBirth() {
        return birth;
    }

    public void setBirth(Date birth) {
        this.birth = birth;
    }

    public Double getMathGrade() {
        return mathGrade;
    }

    public void setMathGrade(Double mathGrade) {
        this.mathGrade = mathGrade;
    }

    public Boolean getVisible() {
        return visible;
    }

    public void setVisible(Boolean visible) {
        this.visible = visible;
    }

    @Override
    public String toString() {
        return "Student{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", sex=" + sex +
                ", birth=" + birth +
                ", mathGrade=" + mathGrade +
                ", visible=" + visible +
                '}';
    }
}
