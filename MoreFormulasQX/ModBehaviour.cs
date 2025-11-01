using Duckov.Economy;
using MoreFormulasQX.Utils;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using UnityEngine;

namespace MoreFormulasQX
{
    public class ModBehaviour : Duckov.Modding.ModBehaviour
    {
        public const string Prefix = nameof(MoreFormulasQX) + ".";

        protected override void OnAfterSetup()
        {
            base.OnAfterSetup();

            HashSet<string> overrideID = new HashSet<string>();
            string filePath = Path.Combine(Application.persistentDataPath, "MoreFormulasCustomConfig.xlsx");
            if (File.Exists(filePath))
            {
                LogHelper.Instance.LogTest("检测到自定义配置文件 MoreFormulasCustomConfig.xlsx，优先加载该文件");
                var overrideformulaInfos = FormulaExcelLoader.Load(filePath);
                foreach (var info in overrideformulaInfos)
                {
                    string formulaID = $"{ModBehaviour.Prefix}{info.formulaID}_formula";
                    overrideID.Add(formulaID);
                    FormulaHelper.AddCraftingFormula(info);
                }
            }

            filePath = null;
            string directoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (directoryName == null) return;
            filePath = Path.Combine(directoryName, "MoreFormulasConfig.xlsx");

            var formulaInfos = FormulaExcelLoader.Load(filePath);

            foreach (var info in formulaInfos)
            {
                string formulaID = $"{ModBehaviour.Prefix}{info.formulaID}_formula";
                if (overrideID.Contains(formulaID))
                {
                    // LogHelper.Instance.LogTest($"配方ID: {formulaID} 存在于自定义配置文件中，跳过加载默认配置");
                    continue;
                }
                FormulaHelper.AddCraftingFormula(info);
            }

            // 添加合成配方
            //Fiber_Lv2();
            //Fiber_Lv3();
            //WeaponPartsLv2();
            //WeaponPartsLv3();
        }

        protected override void OnBeforeDeactivate()
        {
            base.OnBeforeDeactivate();

            // 移除所有增加的合成配方
            FormulaHelper.RemoveAllAddedFormulas();
        }

        /// <summary>
        /// 高级有机纤维
        /// </summary>
        private void Fiber_Lv2()
        {
            // 配方Tag “WorkBenchAdvanced” 
            FormulaHelper.AddCraftingFormula(
                $"{Prefix}{nameof(Fiber_Lv2)}_formula",
                new Cost
                {
                    money = 0,
                    items = new Cost.ItemEntry[]
                    {
                        // 有机纤维(4)
                        new Cost.ItemEntry
                        {
                            id = 743,
                            amount = 4
                        }
                    }
                },
                new global::CraftingFormula.ItemEntry
                {
                    id = 1170,
                    amount = 1
                },
                new string[] { "WorkBenchAdvanced" },
                requirePerk: "",
                unlockByDefault: true,
                hideInIndex: false,
                lockInDemo: false
            );
        }

        /// <summary>
        /// 顶级有机纤维
        /// </summary>
        private void Fiber_Lv3()
        {
            // 配方Tag “WorkBenchAdvanced” 
            FormulaHelper.AddCraftingFormula(
                $"{Prefix}{nameof(Fiber_Lv3)}_formula",
                new Cost
                {
                    money = 0,
                    items = new Cost.ItemEntry[]
                    {
                        // 高级有机纤维(4)
                        new Cost.ItemEntry
                        {
                            id = 1170,
                            amount = 4
                        }
                    }
                },
                new global::CraftingFormula.ItemEntry
                {
                    id = 1171,
                    amount = 1
                },
                new string[] { "WorkBenchAdvanced" },
                requirePerk: "",
                unlockByDefault: true,
                hideInIndex: false,
                lockInDemo: false
            );
        }

        /// <summary>
        /// 中级武器零件
        /// </summary>
        private void WeaponPartsLv2()
        {
            // 配方Tag “WorkBenchAdvanced” 
            FormulaHelper.AddCraftingFormula(
                $"{Prefix}{nameof(WeaponPartsLv2)}_formula",
                new Cost
                {
                    money = 0,
                    items = new Cost.ItemEntry[]
                    {
                        // 武器零件(2)
                        new Cost.ItemEntry
                        {
                            id = 367,
                            amount = 2
                        }
                    }
                },
                new global::CraftingFormula.ItemEntry
                {
                    id = 662,
                    amount = 1
                },
                new string[] { "WorkBenchAdvanced" },
                requirePerk: "",
                unlockByDefault: true,
                hideInIndex: false,
                lockInDemo: false
            );
        }

        /// <summary>
        /// 高级武器零件
        /// </summary>
        private void WeaponPartsLv3()
        {
            // 配方Tag “WorkBenchAdvanced” 
            FormulaHelper.AddCraftingFormula(
                $"{Prefix}{nameof(WeaponPartsLv3)}_formula",
                new Cost
                {
                    money = 0,
                    items = new Cost.ItemEntry[]
                    {
                        // 中级武器零件(4)
                        new Cost.ItemEntry
                        {
                            id = 662,
                            amount = 4
                        }
                    }
                },
                new global::CraftingFormula.ItemEntry
                {
                    id = 663,
                    amount = 1
                },
                new string[] { "WorkBenchAdvanced" },
                requirePerk: "",
                unlockByDefault: true,
                hideInIndex: false,
                lockInDemo: false
            );
        }

    }
}
