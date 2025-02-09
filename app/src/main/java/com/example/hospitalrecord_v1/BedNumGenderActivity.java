package com.example.hospitalrecord_v1;

import androidx.appcompat.app.AppCompatActivity;

import android.content.Context;
import android.content.SharedPreferences;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.RadioButton;
import android.widget.RadioGroup;

public class BedNumGenderActivity extends AppCompatActivity {
    private String TAG = "BedNumGenderActivity";
    private Button btn_bedNumGender_send;
    private EditText edit_bedNumGender_bedNum;
    private RadioGroup rg_bedNumGender_gender;
    private RadioButton rb_bedNumGender_men,rb_bedNumGender_woman;
    private String bedNum,gender;
    SharedPreferences sharedPref;
    SharedPreferences.Editor editor;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_bed_num_gender);
        initView();
    }
    
    private void initView(){
        initUserData();
        edit_bedNumGender_bedNum = (EditText)findViewById(R.id.edit_bedNumGender_bedNum);
        rg_bedNumGender_gender = (RadioGroup) findViewById(R.id.rg_bedNumGender_gender);
        rb_bedNumGender_men = (RadioButton) findViewById(R.id.rb_bedNumGender_men);
        rb_bedNumGender_woman = (RadioButton) findViewById(R.id.rb_bedNumGender_woman);
        btn_bedNumGender_send = (Button) findViewById(R.id.btn_bedNumGender_send);
        if(!bedNum.equals("")){
            edit_bedNumGender_bedNum.setText(bedNum);
        }
        if (gender.equals("男")){
            rb_bedNumGender_men.setChecked(true);
        }else if(gender.equals("女")){
            rb_bedNumGender_woman.setChecked(true);
        }
        rg_bedNumGender_gender.setOnCheckedChangeListener(new RadioGroup.OnCheckedChangeListener() {
            @Override
            public void onCheckedChanged(RadioGroup radioGroup, int checkedId) {
                if (checkedId == R.id.rb_bedNumGender_men){
                    gender = "男";
                }else if (checkedId == R.id.rb_bedNumGender_woman){
                    gender = "女";
                }
            }
        });

        btn_bedNumGender_send.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                Log.d(TAG, "onClick: ");
                setData();
                finish();
            }
        });
        
    }

    private void initUserData(){
        sharedPref = getSharedPreferences("data_storage", Context.MODE_PRIVATE);
        editor = sharedPref.edit();
        bedNum = sharedPref.getString("bedNum","");
        gender = sharedPref.getString("gender","");
        Log.d(TAG, "initUserData: bedNum = "+bedNum);
        Log.d(TAG, "initUserData: gender = "+gender);
    }

    private void setData(){
        bedNum = edit_bedNumGender_bedNum.getText().toString();
        if(!bedNum.equals("")){
            Log.d(TAG, "setData: bedNum");
            Log.d(TAG, "setData: bedNum = "+bedNum);
            editor.putString("bedNum", bedNum);
            editor.apply();
        }
        if (!gender.equals("")){
            Log.d(TAG, "setData: gender");
            editor.putString("gender", gender);
            editor.apply();
        }
    }
    
}