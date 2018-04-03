using FISCA;
using FISCA.Presentation;
using Framework;
using JHSchool;
using JHSchool.Data;

namespace Behavior.TheCadre
{
    public class Program
    {
        //2010/8/18日
        //這是國中所使用的舊版本幹部模組
        //現在已改成國高中共用版本
        //暫時將此程式碼擱置


        [MainMethod("Behavior.TheCadre")]
        static public void Main()
        {

            ServerModule.AutoManaged("http://module.ischool.com.tw/module/138/Cadre_Behavior/udm.xml");

            //班級幹部輸入
            //JHSchool.Class.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<ClassCadreItem>());

            FISCA.Permission.FeatureAce UserPermission;
            UserPermission = FISCA.Permission.UserAcl.Current["Behavior.TheCadre.Detail00040"];
            if (UserPermission.Editable || UserPermission.Viewable)
                JHSchool.Student.Instance.AddDetailBulider(new FISCA.Presentation.DetailBulider<StudentCadreItem>());


            //RibbonBarItem edit = StuAdmin.Instance.RibbonBarItems["幹部管理"];

            //edit["幹部名稱設定"].Enable = User.Acl["Behavior.TheCadre.Ribbon00010"].Executable; 
            //edit["幹部名稱設定"].Click += delegate
            //{
            //    CadreConfig cc = new CadreConfig();
            //    cc.ShowDialog();
            //};

            //edit["學校幹部"].Enable = User.Acl["Behavior.TheCadre.Ribbon00020"].Executable; 
            //edit["學校幹部"].Click += delegate
            //{
            //    SchoolCadreForm SCFform = new SchoolCadreForm();
            //    SCFform.ShowDialog();
            //};

            RibbonBarItem rbItem7 = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"];
            rbItem7["報表"]["學務相關報表"]["班級幹部總表"].Enable = User.Acl["K12.class.TheCadre.Report00060.5"].Executable;
            rbItem7["報表"]["學務相關報表"]["班級幹部總表"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count != 0)
                {
                    StudentLeadersSummaryTable StudentRW = new StudentLeadersSummaryTable();
                    StudentRW.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇班級!");
                }
            };

            RibbonBarItem rbItem8 = FISCA.Presentation.MotherForm.RibbonBarItems["學務作業", "資料統計"];
            rbItem8["報表"].Image = Properties.Resources.paste_64;
            rbItem8["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            rbItem8["報表"]["學校幹部總表"].Enable = User.Acl["K12.class.TheCadre.Report00060.8"].Executable;
            rbItem8["報表"]["學校幹部總表"].Click += delegate
            {
                SchoolLeadersSummaryTable StudentRW = new SchoolLeadersSummaryTable();
                StudentRW.ShowDialog();
            };

            RibbonBarItem rbItem2 = Student.Instance.RibbonBarItems["資料統計"];
            rbItem2["報表"]["學務相關報表"]["學生幹部證明單"].Enable = User.Acl["Behavior.TheCadre.Report00060"].Executable;
            rbItem2["報表"]["學務相關報表"]["學生幹部證明單"].Click += delegate
            {
                if (Student.Instance.SelectedKeys.Count != 0)
                {
                    CadreProveReport StudentRW = new CadreProveReport();
                    StudentRW.ShowDialog();
                }
                else
                {
                    MsgBox.Show("請選擇學生!");
                }
            };



            string URL學生幹部證明單 = "ischool/幹部模組/學生/報表/學務/學生幹部證明單";
            FISCA.Features.Register(URL學生幹部證明單, arg =>
            {
                 CadreProveReport StudentRW = new CadreProveReport();
                 StudentRW.ShowDialog();
            });


            RibbonBarItem rbItem3 = Class.Instance.RibbonBarItems["學務"];
            //rbItem3["班級幹部管理"].Image = Properties.Resources.niche_fav_64;
            //rbItem3["班級幹部管理"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //rbItem3["班級幹部管理"].Enable = false;
            //rbItem3["班級幹部管理"].Click += delegate
            //{
            //    if (Class.Instance.SelectedKeys.Count == 1)
            //    {
            //        TheCadreByClassForm CBC = new TheCadreByClassForm(Class.Instance.SelectedKeys[0]);
            //        CBC.ShowDialog();
            //    }
            //    else if (Class.Instance.SelectedKeys.Count > 1)
            //    {
            //        MsgBox.Show("本功能僅提供對單一班級進行幹部登錄作業!");
            //    }
            //    else
            //    {
            //        MsgBox.Show("請選擇一個班級!!");
            //    }
            //};

            rbItem3["班級幹部登錄"].Enable = false;
            rbItem3["班級幹部登錄"].Image = Properties.Resources.stamp_paper_fav_128;
            rbItem3["班級幹部登錄"].Size = RibbonBarButton.MenuButtonSize.Medium;
            rbItem3["班級幹部登錄"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count == 1)
                {
                    ClassSpeedInsertBySeanNo CBC = new ClassSpeedInsertBySeanNo();
                    CBC.ShowDialog();
                }
                else if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 1)
                {
                    MsgBox.Show("本功能僅提供對單一班級進行幹部登錄作業!");
                }
                else
                {
                    MsgBox.Show("請選擇一個班級!!");
                }
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                rbItem3["班級幹部登錄"].Enable = (User.Acl["Behavior.TheCadre.Report00070.1"].Executable && (K12.Presentation.NLDPanels.Class.SelectedSource.Count == 1));
                //rbItem3["班級幹部管理"].Enable = (User.Acl["Behavior.TheCadre.Report00070"].Executable && (K12.Presentation.NLDPanels.Class.SelectedSource.Count == 1));

                rbItem7["報表"]["學務相關報表"]["班級幹部總表"].Enable = (User.Acl["K12.class.TheCadre.Report00060.5"].Executable && (K12.Presentation.NLDPanels.Class.SelectedSource.Count >= 1));
            };

            #region 匯出及匯入
            RibbonBarButton rbItemImport = Student.Instance.RibbonBarItems["資料統計"]["匯入"];
            RibbonBarButton rbItemExport = Student.Instance.RibbonBarItems["資料統計"]["匯出"];

            rbItemExport["學務相關匯出"]["匯出擔任幹部記錄"].Enable = User.Acl["JHSchool.Student.Ribbon0167"].Executable;
            rbItemExport["學務相關匯出"]["匯出擔任幹部記錄"].Click += delegate
            {
                SmartSchool.API.PlugIn.Export.Exporter exporter = new Behavior.TheCadre.ImportExport.ExportSchoolObject();
                JHSchool.Behavior.ImportExport.ExportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                exporter.InitializeExport(wizard);
                wizard.ShowDialog();
            };

            rbItemImport["學務相關匯入"]["匯入擔任幹部記錄"].Enable = User.Acl["JHSchool.Student.Ribbon0168"].Executable;
            rbItemImport["學務相關匯入"]["匯入擔任幹部記錄"].Click += delegate
            {
                SmartSchool.API.PlugIn.Import.Importer importer = new Behavior.TheCadre.ImportExport.ImportSchoolObject();
                JHSchool.Behavior.ImportExport.ImportStudentV2 wizard = new JHSchool.Behavior.ImportExport.ImportStudentV2(importer.Text, importer.Image);
                importer.InitializeImport(wizard);
                wizard.ShowDialog();
            };

            #endregion

            RibbonBarItem RibbonItem = FISCA.Presentation.MotherForm.RibbonBarItems["學務作業", "基本設定"];
            RibbonItem["管理"]["幹部名稱管理"].Enable = User.Acl["Behavior.TheCadre.Report00070.5"].Executable;
            RibbonItem["管理"]["幹部名稱管理"].Click += delegate
            {
                NewCadreSetup cs1 = new NewCadreSetup();
                cs1.ShowDialog();
            };

            RibbonBarItem RibbonSpeedInsert = FISCA.Presentation.MotherForm.RibbonBarItems["學務作業", "批次作業/查詢"];
            RibbonSpeedInsert["學校幹部登錄"].Enable = User.Acl["Behavior.TheCadre.Report00080"].Executable;
            RibbonSpeedInsert["學校幹部登錄"].Image = Properties.Resources.stamp_paper_fav_128;
            RibbonSpeedInsert["學校幹部登錄"].Click += delegate
            {
                SchoolSpeedInsertByClassSeanNo cs1 = new SchoolSpeedInsertByClassSeanNo();
                cs1.ShowDialog();
            };

            // 2018/03/20 羿均 優化項目
            RibbonSpeedInsert["幹部批次修改"].Enable = User.Acl["E00CFDDF-0F68-46CA-8608-C901A3E10616"].Executable;
            RibbonSpeedInsert["幹部批次修改"].Image = Properties.Resources.niche_fav_64;
            RibbonSpeedInsert["幹部批次修改"].Click += delegate
            {
                (new CardEdit.CadreEditForm()).ShowDialog();
            };

            #region 權限控管
            //Framework.Security.Catalog ribbon = Framework.Security.RoleAclSource.Instance["學務作業"];
            //ribbon.Add(new Framework.Security.RibbonFeature("Behavior.TheCadre.Ribbon00010", "幹部名稱設定"));
            //ribbon.Add(new Framework.Security.RibbonFeature("Behavior.TheCadre.Ribbon00020", "學校幹部"));

            //Framework.Security.Catalog detail = Framework.Security.RoleAclSource.Instance["班級"]["資料項目"];
            //detail.Add(new Framework.Security.DetailItemFeature("Behavior.TheCadre.Detail00030", "班級幹部"));

            Framework.Security.Catalog detail2 = Framework.Security.RoleAclSource.Instance["學生"]["資料項目"];
            detail2.Add(new Framework.Security.DetailItemFeature("Behavior.TheCadre.Detail00040", "幹部記錄"));

            detail2 = Framework.Security.RoleAclSource.Instance["學生"]["報表"];
            detail2.Add(new Framework.Security.ReportFeature("Behavior.TheCadre.Report00060", "學生幹部證明單"));

            detail2 = Framework.Security.RoleAclSource.Instance["學生"]["功能按鈕"];
            detail2.Add(new Framework.Security.RibbonFeature("JHSchool.Student.Ribbon0167", "匯出擔任幹部記錄"));

            detail2 = Framework.Security.RoleAclSource.Instance["學生"]["功能按鈕"];
            detail2.Add(new Framework.Security.RibbonFeature("JHSchool.Student.Ribbon0168", "匯入擔任幹部記錄"));

            detail2 = Framework.Security.RoleAclSource.Instance["班級"]["功能按鈕"];
            //detail2.Add(new Framework.Security.RibbonFeature("Behavior.TheCadre.Report00070", "班級幹部管理"));
            detail2.Add(new Framework.Security.RibbonFeature("Behavior.TheCadre.Report00070.1", "班級幹部登錄"));
            detail2.Add(new Framework.Security.RibbonFeature("K12.class.TheCadre.Report00060.5", "班級幹部總表"));

            detail2 = Framework.Security.RoleAclSource.Instance["學務作業"];
            detail2.Add(new Framework.Security.RibbonFeature("Behavior.TheCadre.Report00070.5", "幹部名稱管理"));
            detail2.Add(new Framework.Security.RibbonFeature("Behavior.TheCadre.Report00080", "學校幹部登錄"));
            detail2.Add(new Framework.Security.RibbonFeature("K12.class.TheCadre.Report00060.8", "學校幹部總表"));
            detail2.Add(new Framework.Security.ReportFeature("E00CFDDF-0F68-46CA-8608-C901A3E10616", "幹部批次修改"));

            #endregion

        }
    }
}