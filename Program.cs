﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;


namespace ExamScoreCardReaderV2
{
    /// <summary>
    /// 成績讀卡
    /// </summary>
    public static class Program
    {
        [MainMethod]
        public static void Main()
        {
            RibbonBarItem rbItem = FISCA.Presentation.MotherForm.RibbonBarItems["課程", "讀卡"];

            RibbonBarButton importButton = rbItem["匯入讀卡成績"];
            importButton.Size = RibbonBarButton.MenuButtonSize.Large;
            importButton.Enable = Permissions.匯入讀卡成績權限;
            importButton.Image = Properties.Resources.byte_add_64;
            importButton.Click += delegate
            {
                ImportStartupForm form = new ImportStartupForm();
                form.ShowDialog();
            };

            // 2016/8/4 穎驊新增試別子成績設定
            RibbonBarButton subscoreButton = rbItem["試別子成績設定"];
            subscoreButton.Size = RibbonBarButton.MenuButtonSize.Large;
            subscoreButton.Image = Properties.Resources.設定;
            subscoreButton.Enable = false;
            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += delegate
            {
                subscoreButton.Enable = Permissions.試別子成績設定權限 & K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0;
            };
            subscoreButton.Click += delegate
            {
                new SubScoreSettingForm().ShowDialog();
            };

            RibbonBarButton classButton = rbItem["班級代碼設定"];
            classButton.Size = RibbonBarButton.MenuButtonSize.Small;
            classButton.Enable = Permissions.班級代碼設定權限;
            classButton.Click += delegate
            {
                new ClassCodeConfig().ShowDialog();
            };

            RibbonBarButton examButton = rbItem["試別代碼設定"];
            examButton.Size = RibbonBarButton.MenuButtonSize.Small;
            examButton.Enable = Permissions.試別代碼設定權限;
            examButton.Click += delegate
            {
                new ExamCodeConfig().ShowDialog();
            };

            RibbonBarButton subjectButton = rbItem["科目代碼設定"];
            subjectButton.Size = RibbonBarButton.MenuButtonSize.Small;
            subjectButton.Enable = Permissions.科目代碼設定權限;
            subjectButton.Click += delegate
            {
                new SubjectCodeConfig().ShowDialog();
            };



            Catalog detail = RoleAclSource.Instance["課程"]["功能按鈕"];
            detail.Add(new ReportFeature(Permissions.匯入讀卡成績, "匯入讀卡成績"));
            detail.Add(new ReportFeature(Permissions.班級代碼設定, "班級代碼設定"));
            detail.Add(new ReportFeature(Permissions.試別代碼設定, "試別代碼設定"));
            detail.Add(new ReportFeature(Permissions.科目代碼設定, "科目代碼設定"));
            detail.Add(new ReportFeature(Permissions.試別子成績設定, "試別子成績設定"));
        }
    }
}
