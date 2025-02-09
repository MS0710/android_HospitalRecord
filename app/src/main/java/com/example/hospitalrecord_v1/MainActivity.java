package com.example.hospitalrecord_v1;

import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.app.Activity;
import android.content.Context;
import android.content.DialogInterface;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.os.Environment;
import android.os.StrictMode;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ImageView;
import android.widget.RadioButton;
import android.widget.RadioGroup;
import android.widget.TextView;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class MainActivity extends AppCompatActivity {
    private String TAG = "";
    private static final int REQUEST_EXTERNAL_STORAGE = 1 ;

    private boolean isPermissionPassed = false;
    private static String[] PERMISSON_STORAGE = {"android.permission.READ_EXTERNAL_STORAGE",
            "android.permission.WRITE_EXTERNAL_STORAGE"};
    WritableWorkbook workbook;
    // 創建工作表
    WritableSheet sheet;
    private ImageView img_main_clear;
    SharedPreferences sharedPref;
    SharedPreferences.Editor editor;
    private String name,date,bedNum,gender;
    private TextView txt_main_bedNum,txt_main_gender,txt_main_name,txt_main_date;
    private Button btn_bedNum_gender,btn_name_date,btn_main_export;
    private RadioGroup[] RG_ID = new RadioGroup[35];
    int[] tempRGID = {R.id.rg_main_AM_washFace, R.id.rg_main_PM_washFace,
            R.id.rg_main_AM_mouth, R.id.rg_main_PM_mouth,
            R.id.rg_main_AM_shower,R.id.rg_main_PM_shower,
            R.id.rg_main_breakfast,R.id.rg_main_lunch,R.id.rg_main_dessert,R.id.rg_main_dinner,
            R.id.rg_main_shit_water,R.id.rg_main_nasogastricTube,R.id.rg_main_urinaryCatheter,
            R.id.rg_main_clothing,R.id.rg_main_sheet,R.id.rg_main_diaper,
            R.id.rg_main_samllDiaper,R.id.rg_main_urineCondom,R.id.rg_main_TrimNails,
            R.id.rg_main_shave,R.id.rg_main_hair,
            R.id.rg_main_rollOver0800,R.id.rg_main_rollOver1000,R.id.rg_main_rollOver1200,
            R.id.rg_main_rollOver1400,R.id.rg_main_rollOver1600,R.id.rg_main_rollOver1800,
            R.id.rg_main_rollOver2000,R.id.rg_main_rollOver2200,R.id.rg_main_rollOver2400,
            R.id.rg_main_rollOver0200,R.id.rg_main_rollOver0400,R.id.rg_main_rollOver0600,
    };
    private RadioButton rb_breakfast_25,rb_breakfast_50,rb_breakfast_75,rb_breakfast_100;
    private RadioButton rb_lunch_25,rb_lunch_50,rb_lunch_75,rb_lunch_100;
    private RadioButton rb_dessert_N,rb_dessert_Y;
    private RadioButton rb_dinner_25,rb_dinner_50,rb_dinner_75,rb_dinner_100;
    private EditText edit_main_shitTime;
    private int shitTimes;
    private RadioButton rb_shit_water,rb_shit_soft,rb_shit_normal,rb_shit_hard;
    private RadioButton rb_nasogastricTube_No,rb_nasogastricTube_notYet,rb_nasogastricTube_done;
    private RadioButton rb_urinaryCatheter_No,rb_urinaryCatheter_notYet,rb_urinaryCatheter_done;
    private EditText edit_main_outBed;
    private int outBrdTimes;
    int[] tempRB_N_ID = {R.id.rb_AM_washFace_N, R.id.rb_PM_washFace_N,
            R.id.rb_AM_mouth_N, R.id.rb_PM_mouth_N,
            R.id.rb_AM_shower_N,R.id.rb_PM_shower_N,
    };
    private RadioButton[] RB_ID_N = new RadioButton[6];
    int[] tempRB_N_ID_P2 = {
            R.id.rb_clothing_N,R.id.rb_sheet_N,R.id.rb_diaper_N,
            R.id.rb_samllDiaper_N,R.id.rb_urineCondom_N,R.id.rb_TrimNails_N,
            R.id.rb_shave_N,R.id.rb_hair_N,
            R.id.rb_rollOver0800_No,R.id.rb_rollOver1000_No,R.id.rb_rollOver1200_No,
            R.id.rb_rollOver1400_No,R.id.rb_rollOver1600_No,R.id.rb_rollOver1800_No,
            R.id.rb_rollOver2000_No,R.id.rb_rollOver2200_No,R.id.rb_rollOver2400_No,
            R.id.rb_rollOver0200_No,R.id.rb_rollOver0400_No,R.id.rb_rollOver0600_No,
    };
    private RadioButton[] RB_ID_N_P2= new RadioButton[20];
    int[] tempRB_Y_ID = {R.id.rb_AM_washFace_Y, R.id.rb_PM_washFace_Y,
            R.id.rb_AM_mouth_Y, R.id.rb_PM_mouth_Y,
            R.id.rb_AM_shower_Y,R.id.rb_PM_shower_Y,
    };
    private RadioButton[] RB_ID_Y= new RadioButton[6];
    int[] tempRB_Y_ID_P2 = {
            R.id.rb_clothing_Y,R.id.rb_sheet_Y,R.id.rb_diaper_Y,
            R.id.rb_samllDiaper_Y,R.id.rb_urineCondom_Y,R.id.rb_TrimNails_Y,
            R.id.rb_shave_Y,R.id.rb_hair_Y,
            R.id.rb_rollOver0800_Yes,R.id.rb_rollOver1000_Yes,R.id.rb_rollOver1200_Yes,
            R.id.rb_rollOver1400_Yes,R.id.rb_rollOver1600_Yes,R.id.rb_rollOver1800_Yes,
            R.id.rb_rollOver2000_Yes,R.id.rb_rollOver2200_Yes,R.id.rb_rollOver2400_Yes,
            R.id.rb_rollOver0200_Yes,R.id.rb_rollOver0400_Yes,R.id.rb_rollOver0600_Yes,
    };
    private RadioButton[] RB_ID_Y_P2= new RadioButton[20];
    private String[] ansFlag = new String[35];
    private TextView txt_main_path;
    private String path;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        path = getDiskCacheDir(MainActivity.this);
        Log.d(TAG, "onCreate: path = "+path);
        initView();
    }

    private void initView(){
        initItem();
        sharedPref = getSharedPreferences("data_storage",Context.MODE_PRIVATE);
        editor = sharedPref.edit();
        initUserData();

        txt_main_bedNum = (TextView) findViewById(R.id.txt_main_bedNum);
        txt_main_gender = (TextView) findViewById(R.id.txt_main_gender);
        txt_main_name = (TextView) findViewById(R.id.txt_main_name);
        txt_main_date = (TextView) findViewById(R.id.txt_main_date);

        for (int i=0; i<tempRGID.length ; i++){
            RG_ID[i] = (RadioGroup) findViewById(tempRGID[i]);
            RG_ID[i].setOnCheckedChangeListener(onCheckChange);
        }
        for (int i=0; i<RB_ID_N.length ; i++){
            RB_ID_N[i] = findViewById(tempRB_N_ID[i]);
            RB_ID_Y[i] = findViewById(tempRB_Y_ID[i]);
        }
        for (int i=0; i<RB_ID_N_P2.length ; i++){
            RB_ID_N_P2[i] = findViewById(tempRB_N_ID_P2[i]);
            RB_ID_Y_P2[i] = findViewById(tempRB_Y_ID_P2[i]);
        }

        rb_breakfast_25 = (RadioButton)findViewById(R.id.rb_breakfast_25);
        rb_breakfast_50 = (RadioButton)findViewById(R.id.rb_breakfast_50);
        rb_breakfast_75 = (RadioButton)findViewById(R.id.rb_breakfast_75);
        rb_breakfast_100 = (RadioButton)findViewById(R.id.rb_breakfast_100);

        rb_lunch_25 = (RadioButton)findViewById(R.id.rb_lunch_25);
        rb_lunch_50 = (RadioButton)findViewById(R.id.rb_lunch_50);
        rb_lunch_75 = (RadioButton)findViewById(R.id.rb_lunch_75);
        rb_lunch_100 = (RadioButton)findViewById(R.id.rb_lunch_100);

        rb_dessert_N = (RadioButton)findViewById(R.id.rb_dessert_N);
        rb_dessert_Y = (RadioButton)findViewById(R.id.rb_dessert_Y);

        rb_dinner_25 = (RadioButton)findViewById(R.id.rb_dinner_25);
        rb_dinner_50 = (RadioButton)findViewById(R.id.rb_dinner_50);
        rb_dinner_75 = (RadioButton)findViewById(R.id.rb_dinner_75);
        rb_dinner_100 = (RadioButton)findViewById(R.id.rb_dinner_100);

        edit_main_shitTime = (EditText)findViewById(R.id.edit_main_shitTime);

        rb_shit_water = (RadioButton)findViewById(R.id.rb_shit_water);
        rb_shit_soft = (RadioButton)findViewById(R.id.rb_shit_soft);
        rb_shit_normal = (RadioButton)findViewById(R.id.rb_shit_normal);
        rb_shit_hard = (RadioButton)findViewById(R.id.rb_shit_hard);

        rb_nasogastricTube_No = (RadioButton)findViewById(R.id.rb_nasogastricTube_No);
        rb_nasogastricTube_notYet = (RadioButton)findViewById(R.id.rb_nasogastricTube_notYet);
        rb_nasogastricTube_done = (RadioButton)findViewById(R.id.rb_nasogastricTube_done);

        rb_urinaryCatheter_No = (RadioButton)findViewById(R.id.rb_urinaryCatheter_No);
        rb_urinaryCatheter_notYet = (RadioButton)findViewById(R.id.rb_urinaryCatheter_notYet);
        rb_urinaryCatheter_done = (RadioButton)findViewById(R.id.rb_urinaryCatheter_done);

        edit_main_outBed = (EditText)findViewById(R.id.edit_main_outBed);

        img_main_clear = (ImageView) findViewById(R.id.img_main_clear);
        btn_main_export = (Button) findViewById(R.id.btn_main_export);
        btn_bedNum_gender = (Button) findViewById(R.id.btn_bedNum_gender);
        btn_name_date = (Button) findViewById(R.id.btn_name_date);
        txt_main_path = (TextView)findViewById(R.id.txt_main_path);
        btn_main_export.setOnClickListener(onClick);
        btn_bedNum_gender.setOnClickListener(onClick);
        btn_name_date.setOnClickListener(onClick);
        img_main_clear.setOnClickListener(onClick);

    }

    @Override
    protected void onResume() {
        super.onResume();
        Log.d(TAG, "onResume: ");
        initUserData();
        setUserText();
    }

    private void initItem(){
        for (int i=0; i<35; i++){
            ansFlag[i] = "";
        }
        shitTimes = 0;
        outBrdTimes = 0;
    }

    private void resetAllItem(){
        edit_main_shitTime.setText("");
        edit_main_outBed.setText("");
        for (int i=0; i<RB_ID_N.length ; i++){
            RB_ID_N[i].setChecked(false);
            RB_ID_Y[i].setChecked(false);
        }
        rb_breakfast_25.setChecked(false);
        rb_breakfast_50.setChecked(false);
        rb_breakfast_75.setChecked(false);
        rb_breakfast_100.setChecked(false);
        rb_lunch_25.setChecked(false);
        rb_lunch_50.setChecked(false);
        rb_lunch_75.setChecked(false);
        rb_lunch_100.setChecked(false);
        rb_dessert_N.setChecked(false);
        rb_dessert_Y.setChecked(false);
        rb_dinner_25.setChecked(false);
        rb_dinner_50.setChecked(false);
        rb_dinner_75.setChecked(false);
        rb_dinner_100.setChecked(false);
        rb_shit_water.setChecked(false);
        rb_shit_soft.setChecked(false);
        rb_shit_normal.setChecked(false);
        rb_shit_hard.setChecked(false);
        rb_nasogastricTube_No.setChecked(false);
        rb_nasogastricTube_notYet.setChecked(false);
        rb_nasogastricTube_done.setChecked(false);
        rb_urinaryCatheter_No.setChecked(false);
        rb_urinaryCatheter_notYet.setChecked(false);
        rb_urinaryCatheter_done.setChecked(false);
        for (int i=0; i<RB_ID_N_P2.length ; i++){
            RB_ID_N_P2[i].setChecked(false);
            RB_ID_Y_P2[i].setChecked(false);
        }
    }
    private void initUserData(){
        name = sharedPref.getString("name","");
        date = sharedPref.getString("date","");
        bedNum = sharedPref.getString("bedNum","");
        gender = sharedPref.getString("gender","");
        Log.d(TAG, "initUserData: name = "+name);
        Log.d(TAG, "initUserData: date = "+date);
        Log.d(TAG, "initUserData: bedNum = "+bedNum);
        Log.d(TAG, "initUserData: gender = "+gender);
    }

    private void setUserText(){
        Log.d(TAG, "onClick setUserText: ");
        if(name.equals("")){
            Log.d(TAG, "onClick setUserText: 1");
            txt_main_name.setText("未輸入");
        }else {
            Log.d(TAG, "onClick setUserText: 2");
            txt_main_name.setText(name);
        }

        if(date.equals("")){
            txt_main_date.setText("未輸入");
        }else {
            txt_main_date.setText(date);
        }

        if(bedNum.equals("")){
            Log.d(TAG, "onClick setUserText: 3");
            txt_main_bedNum.setText("未輸入");
        }else {
            Log.d(TAG, "onClick setUserText: 4");
            txt_main_bedNum.setText(bedNum);
        }

        if(gender.equals("")){
            txt_main_gender.setText("未輸入");
        }else {
            txt_main_gender.setText(gender);
        }
    }

    private View.OnClickListener onClick = new View.OnClickListener() {
        @Override
        public void onClick(View view) {
            if (view.getId() == R.id.btn_main_export){
                if(!edit_main_shitTime.getText().toString().isEmpty()){
                    shitTimes = Integer.parseInt(edit_main_shitTime.getText().toString());
                    ansFlag[10] = ""+shitTimes;
                }
                if(!edit_main_outBed.getText().toString().isEmpty()){
                    outBrdTimes = Integer.parseInt(edit_main_outBed.getText().toString());
                    ansFlag[34] = ""+outBrdTimes;
                }
                createExcel();


            }else if (view.getId() == R.id.btn_bedNum_gender){
                Log.d(TAG, "onClick: btn_bedNum_gender");
                Intent intent = new Intent(MainActivity.this, BedNumGenderActivity.class);
                startActivity(intent);
            }else if (view.getId() == R.id.btn_name_date){
                Log.d(TAG, "onClick: btn_name_date");
                Intent intent = new Intent(MainActivity.this, NameDateActivity.class);
                startActivity(intent);
            }else if (view.getId() == R.id.img_main_clear){
                Log.d(TAG, "onClick: img_main_clear");
                AlertDialog.Builder alertDialog = new AlertDialog.Builder(MainActivity.this);
                alertDialog.setTitle("清除資料");
                alertDialog.setMessage("請問要清除當前資料嗎?");
                alertDialog.setPositiveButton("確定", new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface dialog, int which) {
                        Log.d(TAG, "onClick: 確定");
                        editor.clear();
                        editor.commit();
                        initUserData();
                        setUserText();
                        resetAllItem();
                    }
                });
                alertDialog.setNeutralButton("取消",(dialog, which) -> {
                    Log.d(TAG, "onClick: 取消");
                });
                alertDialog.setCancelable(false);
                alertDialog.show();
            }
        }
    };

    private RadioGroup.OnCheckedChangeListener onCheckChange = new RadioGroup.OnCheckedChangeListener() {
        @Override
        public void onCheckedChanged(RadioGroup radioGroup, int i) {
            for (int x=0;x<6;x++){
                if(radioGroup.getId() == tempRGID[x]){
                    if (i == tempRB_N_ID[x]){
                        ansFlag[x] = "No";
                    }
                    if (i == tempRB_Y_ID[x]){
                        ansFlag[x] = "Yes";
                    }
                }
            }
            int j;
            j = 6;
            if(radioGroup.getId() == tempRGID[j]){
                if (i == R.id.rb_breakfast_25){
                    ansFlag[j] = rb_breakfast_25.getText().toString();
                }else if (i == R.id.rb_breakfast_50){
                    ansFlag[j] = rb_breakfast_50.getText().toString();
                }else if (i == R.id.rb_breakfast_75){
                    ansFlag[j] = rb_breakfast_75.getText().toString();
                }else if (i == R.id.rb_breakfast_100){
                    ansFlag[j] = rb_breakfast_100.getText().toString();
                }
            }
            j = 7;
            if(radioGroup.getId() == tempRGID[j]){
                if (i == R.id.rb_lunch_25){
                    ansFlag[j] = rb_lunch_25.getText().toString();
                }else if (i == R.id.rb_lunch_50){
                    ansFlag[j] = rb_lunch_50.getText().toString();
                }else if (i == R.id.rb_lunch_75){
                    ansFlag[j] = rb_lunch_75.getText().toString();
                }else if (i == R.id.rb_lunch_100){
                    ansFlag[j] = rb_lunch_100.getText().toString();
                }
            }
            j = 8;
            if(radioGroup.getId() == tempRGID[j]){
                if (i == R.id.rb_dessert_N){
                    ansFlag[j] = rb_dessert_N.getText().toString();
                }else if (i == R.id.rb_dessert_Y){
                    ansFlag[j] = rb_dessert_Y.getText().toString();
                }
            }
            j = 9;
            if(radioGroup.getId() == tempRGID[j]){
                if (i == R.id.rb_dinner_25){
                    ansFlag[j] = rb_dinner_25.getText().toString();
                }else if (i == R.id.rb_dinner_50){
                    ansFlag[j] = rb_dinner_50.getText().toString();
                }else if (i == R.id.rb_dinner_75){
                    ansFlag[j] = rb_dinner_75.getText().toString();
                }else if (i == R.id.rb_dinner_100){
                    ansFlag[j] = rb_dinner_100.getText().toString();
                }
            }
            j = 11;
            if(radioGroup.getId() == tempRGID[10]){
                if (i == R.id.rb_shit_water){
                    ansFlag[j] = rb_shit_water.getText().toString();
                }else if (i == R.id.rb_shit_soft){
                    ansFlag[j] = rb_shit_soft.getText().toString();
                }else if (i == R.id.rb_shit_normal){
                    ansFlag[j] = rb_shit_normal.getText().toString();
                }else if (i == R.id.rb_shit_hard){
                    ansFlag[j] = rb_shit_hard.getText().toString();
                }
            }
            j = 12;
            if(radioGroup.getId() == tempRGID[11]){
                if (i == R.id.rb_nasogastricTube_No){
                    ansFlag[j] = rb_nasogastricTube_No.getText().toString();
                }else if (i == R.id.rb_nasogastricTube_notYet){
                    ansFlag[j] = rb_nasogastricTube_notYet.getText().toString();
                }else if (i == R.id.rb_nasogastricTube_done){
                    ansFlag[j] = rb_nasogastricTube_done.getText().toString();
                }
            }
            j = 13;
            if(radioGroup.getId() == tempRGID[12]){
                if (i == R.id.rb_urinaryCatheter_No){
                    ansFlag[j] = rb_urinaryCatheter_No.getText().toString();
                }else if (i == R.id.rb_urinaryCatheter_notYet){
                    ansFlag[j] = rb_urinaryCatheter_notYet.getText().toString();
                }else if (i == R.id.rb_urinaryCatheter_done){
                    ansFlag[j] = rb_urinaryCatheter_done.getText().toString();
                }
            }
            for (int x=0 ;x<20;x++){
                if(radioGroup.getId() == tempRGID[13+x]){
                    if (i == tempRB_N_ID_P2[x]){
                        ansFlag[14+x] = "No";
                    }else if (i == tempRB_Y_ID_P2[x]){
                        ansFlag[14+x] = "Yes";
                    }
                }
            }
        }
    };

    public String getDiskCacheDir(Context context) {
        String cachePath = null;
        if (Environment.MEDIA_MOUNTED.equals(Environment.getExternalStorageState())
                || !Environment.isExternalStorageRemovable()) {
            cachePath = context.getExternalCacheDir().getPath();
        } else {
            cachePath = context.getCacheDir().getPath();
        }
        return cachePath;
    }

    private void createExcel(){
        String mPath = path +"/file_"+date+"_"+name+".xls";
        txt_main_path.setText(mPath);
        Log.d(TAG, "createExcel: path = "+mPath);
        try {
            workbook = Workbook.createWorkbook(new File(mPath));
            sheet = workbook.createSheet("Sheet1", 0);

            Label header1 = new Label(0, 0, "床號: "+bedNum);
            Label header2 = new Label(1, 0, "性別: "+gender);
            Label header3 = new Label(0, 1, "日期: "+date);
            Label header4 = new Label(1, 1, "姓名: "+name);
            sheet.addCell(header1);
            sheet.addCell(header2);
            sheet.addCell(header3);
            sheet.addCell(header4);

            Label[] itemTitle = new Label[35];
            Label[] itemResult = new Label[35];
            itemTitle[0] = new Label(0, 2, "白班洗臉");
            itemTitle[1] = new Label(0, 3, "晚班洗臉");
            itemTitle[2] = new Label(0, 4, "白班口腔");
            itemTitle[3] = new Label(0, 5, "晚班口腔");
            itemTitle[4] = new Label(0, 6, "白班洗浴");
            itemTitle[5] = new Label(0, 7, "晚班洗浴");
            itemTitle[6] = new Label(0, 8, "早餐(管灌)");
            itemTitle[7] = new Label(0, 9, "午餐(管灌)");
            itemTitle[8] = new Label(0, 10, "點心(管灌)");
            itemTitle[9] = new Label(0, 11, "晚餐(管灌)");
            itemTitle[10] = new Label(0, 12, "大便次數");
            itemTitle[11] = new Label(0, 13, "大便性質");
            itemTitle[12] = new Label(0, 14, "鼻胃管護理");
            itemTitle[13] = new Label(0, 15, "尿管護理");
            itemTitle[14] = new Label(0, 16, "換衣服");
            itemTitle[15] = new Label(0, 17, "換床單");
            itemTitle[16] = new Label(0, 18, "換尿布");
            itemTitle[17] = new Label(0, 19, "換小尿片");
            itemTitle[18] = new Label(0, 20, "換尿套");
            itemTitle[19] = new Label(0, 21, "剪指甲");
            itemTitle[20] = new Label(0, 22, "刮鬍子");
            itemTitle[21] = new Label(0, 23, "整理頭髮");
            itemTitle[22] = new Label(0, 24, "翻身08:00");
            itemTitle[23] = new Label(0, 25, "翻身10:00");
            itemTitle[24] = new Label(0, 26, "翻身12:00");
            itemTitle[25] = new Label(0, 27, "翻身14:00");
            itemTitle[26] = new Label(0, 28, "翻身16:00");
            itemTitle[27] = new Label(0, 29, "翻身18:00");
            itemTitle[28] = new Label(0, 30, "翻身20:00");
            itemTitle[29] = new Label(0, 31, "翻身22:00");
            itemTitle[30] = new Label(0, 32, "翻身24:00");
            itemTitle[31] = new Label(0, 33, "翻身02:00");
            itemTitle[32] = new Label(0, 34, "翻身04:00");
            itemTitle[33] = new Label(0, 35, "翻身06:00");
            itemTitle[34] = new Label(0, 36, "下床次數");

            for (int i=0;i<35;i++){
                itemResult[i] = new Label(1,2+i, ansFlag[i]);
            }

            for (int i=0;i<35;i++){
                sheet.addCell(itemTitle[i]);
            }
            for (int i=0;i<35;i++){
                sheet.addCell(itemResult[i]);
            }

            WritableCellFormat headerFormat = new WritableCellFormat();
            headerFormat.setAlignment(Alignment.CENTRE);
            header1.setCellFormat(headerFormat);
            header2.setCellFormat(headerFormat);
            header3.setCellFormat(headerFormat);
            header4.setCellFormat(headerFormat);

            WritableCellFormat dataFormat = new WritableCellFormat();
            dataFormat.setAlignment(Alignment.CENTRE);
            for (int i=0;i<35;i++){
                itemTitle[i].setCellFormat(dataFormat);
            }
            for (int i=0;i<35;i++){
                itemResult[i].setCellFormat(dataFormat);
            }

            workbook.write();
            workbook.close();

        }catch (Exception e){
            e.printStackTrace();
        }

    }


}