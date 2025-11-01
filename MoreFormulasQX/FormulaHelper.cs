using Duckov.Economy;
using MoreFormulasQX.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using UnityEngine;

namespace MoreFormulasQX
{
    public class FormulaHelper
    {
        private static HashSet<string> addedFormulaIDs = new HashSet<string>();

        public static void AddCraftingFormula(FormulaExcelLoader.CraftingFormulaInfo craftingFormulaInfo, string requirePerk = "", bool unlockByDefault = true, bool hideInIndex = false, bool lockInDemo = false)
        {
            if (!craftingFormulaInfo.enabled) return;
            string formulaID = $"{ModBehaviour.Prefix}{craftingFormulaInfo.formulaID}_formula";
            AddCraftingFormula(
                formulaID,
                craftingFormulaInfo.cost,
                craftingFormulaInfo.resultItem,
                craftingFormulaInfo.tags,
                requirePerk,
                unlockByDefault,
                hideInIndex,
                lockInDemo
            );
            Debug.LogWarning($"物品配方：{formulaID} 已添加");
        }


        public static void AddCraftingFormula(string formulaID, Cost costInfo, global::CraftingFormula.ItemEntry resultItemInfo, string[] tags = null, string requirePerk = "", bool unlockByDefault = true, bool hideInIndex = false, bool lockInDemo = false)
        {
            try
            {
                CraftingFormulaCollection instance = CraftingFormulaCollection.Instance;
                // 获取配方列表
                List<global::CraftingFormula> list = ReflectionHelper.GetFieldValue<List<global::CraftingFormula>>(instance, "list");
                if (list.Any((craftingFormula) => craftingFormula.id == formulaID))
                {
                    Debug.LogWarning($"配方ID: {formulaID} 已存在，跳过添加");
                    return;
                }

                if (tags == null)
                {
                    tags = new string[] { "WorkBenchAdvanced" };
                }

                global::CraftingFormula craftingFormula = new global::CraftingFormula
                {
                    id = formulaID,
                    unlockByDefault = unlockByDefault,
                    cost = costInfo,
                    result = resultItemInfo,
                    requirePerk = requirePerk,
                    tags = tags,
                    hideInIndex = hideInIndex,
                    lockInDemo = lockInDemo
                };

                list.Add(craftingFormula);
                addedFormulaIDs.Add(formulaID);
                ReflectionHelper.SetFieldValue(instance, "_entries_ReadOnly", null);
            }
            catch (Exception ex)
            {
                Debug.LogError($"添加合成配方失败: {ex.Message}");
            }
        }

        public static void RemoveAllAddedFormulas()
        {
            try
            {
                CraftingFormulaCollection instance = CraftingFormulaCollection.Instance;
                // 获取配方列表
                List<global::CraftingFormula> list = ReflectionHelper.GetFieldValue<List<global::CraftingFormula>>(instance, "list");
                list.RemoveAll(craftingFormula => addedFormulaIDs.Contains(craftingFormula.id));

                addedFormulaIDs.Clear();
                ReflectionHelper.SetFieldValue(instance, "_entries_ReadOnly", null);
            }
            catch (Exception ex)
            {
                Debug.LogError($"移除配方时出错: {ex.Message}");
            }
        }
    }
}
