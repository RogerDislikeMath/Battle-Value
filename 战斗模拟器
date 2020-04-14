# 战斗模拟器
## 一、前言  
### 战斗模拟器集成了游戏中大部分战斗系统数值，包括如下几个部分，以下针对Sheet页签进行说明。

      1.关卡vs-循环用调用其他sheet页关键数值信息。
      2.关卡英雄属性new-所有关卡的预估战力数值。
      3.关卡-怪物属性-每个关卡怪物配置及需要提前设定的战斗难度参数。
      4.怪物属性-战斗模拟器输出结果1，对应v4monstertype表。
      5.quest表-怪物站位信息，对应V4.ShowWaveInfo表。
      6.quest2-关卡怪物头像信息，对应多张关卡表。
      7.questnew-开辟vba可视化内存。
      8.MonsterTrait-负责模拟器关卡怪物配置数据有效性，怪物数值特点。
      9.怪物资源索引表-就是名字的意思啦。
      
## 二、整体设计思路

      1.设计之初考虑了战斗节奏战斗时长等问题，经过多个版本的测试将时长定为6-8回合（每个玩家平均出手回合）。
      2.设计之初考虑了诸多预留问题，将战斗波次约定为不大于3波。
      3.使用者需要综合考虑经济数值，根据预估在HotHeroAttr中选出关卡推荐战力，给出关键卡点玩家战斗数值，进而生成每个关卡的玩家数值。
      4.设置怪物数值时要根据技能表给出每回合平均伤害系数。
      5.设置每波怪物消灭回合，怪物平均防御，波次损失生命比例。
      6.经过以上参数可计算出每关怪物的所有平均属性。
      7.使用VBA循环将各关卡的平均怪物属性进行怪物特点放缩，并填在关卡-怪物属性中。
      8.根据已经填好的关卡-怪物属性表，生成v4monstertype表，关卡站位表，怪物头像表。

## 三、代码含义

### 循环，将怪物数值个性化，并填入关卡怪物属性表中

```markdown
Sub 生成怪物数值()
For 关卡序数 = 3 To 675 Step 3   '3 675 对应着关卡-怪物属性的怪物数值的首尾行
Sheets("关卡vs").Cells(1, 2) = Sheets("关卡-怪物属性").Cells(关卡序数, 1)         '难度列转化到关卡vs表
Sheets("关卡vs").Cells(2, 2) = Sheets("关卡-怪物属性").Cells(关卡序数, 2)         '岛屿列转化到关卡vs表
Sheets("关卡vs").Cells(2, 1) = CInt(Sheets("关卡-怪物属性").Cells(关卡序数, 3))   '关卡列转化到关卡vs表
For 波数 = 1 To 3
Sheets("关卡vs").Cells(1, 4) = 波数   '波次列转化到关卡vs表
For 每波怪物 = 1 To 6
If Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 每波怪物 + 4) = "/" Then
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 19) = ""
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 20) = ""
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 21) = ""
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 22) = ""
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 23) = ""
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 24) = ""
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 25) = ""
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 26) = ""
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 27) = ""
Else
For 怪物特点 = 3 To 60    '3 60 对应着MonsterTrait的怪物数值的首尾行
If Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 每波怪物 + 4) = Sheets("MonsterTrait").Cells(怪物特点, 2) Then
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 19) = CLng(Sheets("关卡vs").Cells(9, 5) * Sheets("MonsterTrait").Cells(怪物特点, 3) / 100)
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 20) = CLng(Sheets("关卡vs").Cells(9, 6) * Sheets("MonsterTrait").Cells(怪物特点, 4) / 100)
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 21) = CLng(Sheets("关卡vs").Cells(9, 7) + Sheets("MonsterTrait").Cells(怪物特点, 5))
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 22) = CLng(Sheets("关卡vs").Cells(9, 8) * Sheets("MonsterTrait").Cells(怪物特点, 6) / 100)
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 23) = CLng(Sheets("关卡vs").Cells(9, 9) + Sheets("MonsterTrait").Cells(怪物特点, 7) / 100)
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 24) = CLng(Sheets("关卡vs").Cells(9, 10) * 0.98 * Sheets("MonsterTrait").Cells(怪物特点, 8) / 100)
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 25) = CLng(Sheets("关卡vs").Cells(9, 11) + Sheets("MonsterTrait").Cells(怪物特点, 9))
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 26) = CLng(Sheets("关卡vs").Cells(9, 12) + Sheets("MonsterTrait").Cells(怪物特点, 10))
Sheets("关卡-怪物属性").Cells(关卡序数 + 波数 - 1, 9 * (每波怪物 - 1) + 27) = CLng(Sheets("关卡vs").Cells(9, 13) + Sheets("MonsterTrait").Cells(怪物特点, 11))
Else
End If
Next 怪物特点
End If
Next 每波怪物
Next 波数
Next 关卡序数
End Sub
```

### 循环，将关卡-怪物属性表中的怪物数值生成怪物编号并填入怪物属性中

```markdown
Sub 生成monstertype()
C = 1
For 关卡序数 = 3 To 675 Step 3 '3 675 对应着关卡-怪物属性的怪物数值的首尾行
d = 0
For A = 1 To 3
For B = 1 To 6
If Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, B + 4) = "/" Then
Else
C = C + 1
d = d + 1
Sheets("怪物属性").Cells(C, 1) = 100000 * NanDu(Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 1)) + 10000 * DaoYu(Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 2)) + 100 * Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 3) + d
Sheets("怪物属性").Cells(C, 2) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 4 + B)
Sheets("怪物属性").Cells(C, 9) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 19 + (B - 1) * 9)
Sheets("怪物属性").Cells(C, 10) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 20 + (B - 1) * 9)
Sheets("怪物属性").Cells(C, 11) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 21 + (B - 1) * 9)
Sheets("怪物属性").Cells(C, 12) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 22 + (B - 1) * 9)
Sheets("怪物属性").Cells(C, 13) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 23 + (B - 1) * 9)
Sheets("怪物属性").Cells(C, 14) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 24 + (B - 1) * 9)
Sheets("怪物属性").Cells(C, 15) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 25 + (B - 1) * 9)
Sheets("怪物属性").Cells(C, 16) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 26 + (B - 1) * 9)
Sheets("怪物属性").Cells(C, 17) = Sheets("关卡-怪物属性").Cells(关卡序数 + A - 1, 27 + (B - 1) * 9)

End If
Next B
Next A
Next 关卡序数
End Sub


```



