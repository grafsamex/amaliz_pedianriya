import pandas as pd
import os
import time
from slovar import spec_ped, gastro,vmp_ped, nefro, gemat, revmo, nevro1, nevro2, orfan, dermat, allerg


def filecsvlist(file: list) -> list:
    """Функция принимает все названия файлов и папок, находящихся в корневой с программой и возвращает список из файлов,
    относящихся к файлам баз данных с расширением .CSV"""
    listdbffiles = list()
    for i in file:
        if i.endswith('.CSV') or i.endswith('.csv') or i.endswith('.xlsx') or i.endswith('.XLSX'):
            listdbffiles.append(i)
    return listdbffiles

def org_tabl(org):
    tabl = pd.DataFrame(columns=['Наименование медицинская организация', 'Профиль койки', 'Количество коек',
                                 'По МО без МТР Случаи', 'По МО без МТР Дни. посещ.', 'По МО без МТР Стоимость(руб)',
                                 'По межтерр помощи МТР Случаи', 'По межтерр помощи МТР Дни. посещ.',
                                 'По межтерр помощи МТР Стоимость(руб)', 'Всего Случаи', 'Всего Дни. посещ.',
                                 'Всего Стоимость(руб)', 'Нормативный показатель работы койки', 'Работа койки',
                                 'Расчетное значение коек', 'Профиль помощи', '!По МО без МТР Случаи',
                                 '!По МО без МТР Дни. посещ.', '!По МО без МТР Стоимость(руб)',
                                 '!По межтерр помощи МТР Случаи', '!По межтерр помощи МТР Дни. посещ.',
                                 '!По межтерр помощи МТР Стоимость(руб)', '!Всего Случаи', '!Всего Дни. посещ.',
                                 '!Всего Стоимость(руб)', '!Расчетное значение коек'])
    data_org = df.loc[(df['4'] == org)]
    spec_df = data_org[data_org['11'].isin(spec_ped)]
    vmp_df = data_org[data_org['11'].isin(vmp_ped)]
    gastro_df = data_org[data_org['13'].isin(gastro)]
    nefro_df = data_org[data_org['13'].isin(nefro)]
    gemat_df = data_org[data_org['13'].isin(gemat)]
    revmo_df = data_org[data_org['13'].isin(revmo)]
    nevro1_df = data_org[data_org['13'].isin(nevro1)]
    nevro2_df = data_org[data_org['13'].isin(nevro2)]
    orfan_df = data_org[data_org['13'].isin(orfan)]
    dermat_df = data_org[data_org['13'].isin(dermat)]
    allerg_df = data_org[data_org['13'].isin(allerg)]
    new_values1 = [org, 'педиатрии специализированная', '', spec_df['15'].sum(), spec_df['16'].sum(),
                  spec_df['17'].sum(), spec_df['18'].sum(), spec_df['19'].sum(), spec_df['20'].sum(),
                  spec_df['15'].sum()+spec_df['18'].sum(), spec_df['16'].sum()+spec_df['19'].sum(),
                  spec_df['17'].sum() + spec_df['20'].sum(), 335, 0, (spec_df['16'].sum()+spec_df['19'].sum()) / 335,
                  'гастроэнтерология', gastro_df['15'].sum(), gastro_df['16'].sum(),
                  gastro_df['17'].sum(), gastro_df['18'].sum(), gastro_df['19'].sum(), gastro_df['20'].sum(),
                  gastro_df['15'].sum()+ gastro_df['18'].sum(), gastro_df['16'].sum() + gastro_df['19'].sum(),
                  gastro_df['17'].sum() + gastro_df['20'].sum(), (gastro_df['16'].sum()+gastro_df['19'].sum()) / 335]
    new_values2 = [org, 'педиатрии вмп', '', vmp_df['15'].sum(), vmp_df['16'].sum(),
                   vmp_df['17'].sum(), vmp_df['18'].sum(), vmp_df['19'].sum(), vmp_df['20'].sum(),
                   vmp_df['15'].sum() + vmp_df['18'].sum(), vmp_df['16'].sum() + vmp_df['19'].sum(),
                   vmp_df['17'].sum() + vmp_df['20'].sum(), 335, 0, (vmp_df['16'].sum() + vmp_df['19'].sum()) / 335,
                   'нефрология', nefro_df['15'].sum(), nefro_df['16'].sum(),
                   nefro_df['17'].sum(), nefro_df['18'].sum(), nefro_df['19'].sum(), nefro_df['20'].sum(),
                   nefro_df['15'].sum() + nefro_df['18'].sum(), nefro_df['16'].sum() + nefro_df['19'].sum(),
                   nefro_df['17'].sum() + nefro_df['20'].sum(), (nefro_df['16'].sum() + nefro_df['19'].sum()) / 335]
    new_values3 = [org, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'гематология', gemat_df['15'].sum(), gemat_df['16'].sum(),
                   gemat_df['17'].sum(), gemat_df['18'].sum(), gemat_df['19'].sum(), gemat_df['20'].sum(),
                   gemat_df['15'].sum() + gemat_df['18'].sum(), gemat_df['16'].sum() + gemat_df['19'].sum(),
                   gemat_df['17'].sum() + gemat_df['20'].sum(), (gemat_df['16'].sum() + gemat_df['19'].sum()) / 335]
    new_values4 = [org, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'ревматология', revmo_df['15'].sum(),
                   revmo_df['16'].sum(),
                   revmo_df['17'].sum(), revmo_df['18'].sum(), revmo_df['19'].sum(), revmo_df['20'].sum(),
                   revmo_df['15'].sum() + revmo_df['18'].sum(), revmo_df['16'].sum() + revmo_df['19'].sum(),
                   revmo_df['17'].sum() + revmo_df['20'].sum(), (revmo_df['16'].sum() + revmo_df['19'].sum()) / 335]
    new_values5 = [org, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'неврология1(ЭПИ)', nevro1_df['15'].sum(),
                   nevro1_df['16'].sum(),
                   nevro1_df['17'].sum(), nevro1_df['18'].sum(), nevro1_df['19'].sum(), nevro1_df['20'].sum(),
                   nevro1_df['15'].sum() + nevro1_df['18'].sum(), nevro1_df['16'].sum() + nevro1_df['19'].sum(),
                   nevro1_df['17'].sum() + nevro1_df['20'].sum(), (nevro1_df['16'].sum() + nevro1_df['19'].sum()) / 335]
    new_values6 = [org, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'неврология2', nevro2_df['15'].sum(),
                   nevro2_df['16'].sum(),
                   nevro2_df['17'].sum(), nevro2_df['18'].sum(), nevro2_df['19'].sum(), nevro2_df['20'].sum(),
                   nevro2_df['15'].sum() + nevro2_df['18'].sum(), nevro2_df['16'].sum() + nevro2_df['19'].sum(),
                   nevro2_df['17'].sum() + nevro2_df['20'].sum(), (nevro2_df['16'].sum() + nevro2_df['19'].sum()) / 335]
    new_values7 = [org, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'орфанные', orfan_df['15'].sum(),
                   orfan_df['16'].sum(),
                   orfan_df['17'].sum(), orfan_df['18'].sum(), orfan_df['19'].sum(), orfan_df['20'].sum(),
                   orfan_df['15'].sum() + orfan_df['18'].sum(), orfan_df['16'].sum() + orfan_df['19'].sum(),
                   orfan_df['17'].sum() + orfan_df['20'].sum(), (orfan_df['16'].sum() + orfan_df['19'].sum()) / 335]
    new_values8 = [org, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'дерматология', dermat_df['15'].sum(),
                   dermat_df['16'].sum(),
                   dermat_df['17'].sum(), dermat_df['18'].sum(), dermat_df['19'].sum(), dermat_df['20'].sum(),
                   dermat_df['15'].sum() + dermat_df['18'].sum(), dermat_df['16'].sum() + dermat_df['19'].sum(),
                   dermat_df['17'].sum() + dermat_df['20'].sum(), (dermat_df['16'].sum() + dermat_df['19'].sum()) / 335]
    new_values9 = [org, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'аллергология и иммунология', allerg_df['15'].sum(),
                   allerg_df['16'].sum(),
                   allerg_df['17'].sum(), allerg_df['18'].sum(), allerg_df['19'].sum(), allerg_df['20'].sum(),
                   allerg_df['15'].sum() + allerg_df['18'].sum(), allerg_df['16'].sum() + allerg_df['19'].sum(),
                   allerg_df['17'].sum() + allerg_df['20'].sum(), (allerg_df['16'].sum() + allerg_df['19'].sum()) / 335]
    new_values10 = ['Итого, ' + org, '', '', new_values1[3] + new_values2[3], new_values1[4] + new_values2[4],
                    new_values1[5] + new_values2[5], new_values1[6] + new_values2[6], new_values1[7] + new_values2[7],
                    new_values1[8] + new_values2[8], new_values1[9] + new_values2[9], new_values1[10] + new_values2[10],
                    new_values1[11] + new_values2[11], new_values1[12] + new_values2[12], new_values1[13] + new_values2[13],
                    new_values1[14] + new_values2[14], '', new_values1[16] + new_values2[16] + new_values3[16]+
                    + new_values4[16] + new_values5[16] + new_values6[16] + new_values7[16] + new_values8[16] + new_values9[16],
                    new_values1[17] + new_values2[17] + new_values3[17] + new_values4[17] + new_values5[17] +
                    + new_values6[17] + new_values7[17] + new_values8[17] + new_values9[17],
                    new_values1[18] + new_values2[18] + new_values3[18] + new_values4[18] + new_values5[18] +
                    + new_values6[18] + new_values7[18] + new_values8[18] + new_values9[18],
                    new_values1[19] + new_values2[19] + new_values3[19] + new_values4[19] + new_values5[19] +
                    + new_values6[19] + new_values7[19] + new_values8[19] + new_values9[19],
                    new_values1[20] + new_values2[20] + new_values3[20] + new_values4[20] + new_values5[20] +
                    + new_values6[20] + new_values7[20] + new_values8[20] + new_values9[20],
                    new_values1[21] + new_values2[21] + new_values3[21] + new_values4[21] + new_values5[21] +
                    + new_values6[21] + new_values7[21] + new_values8[21] + new_values9[21],
                    new_values1[22] + new_values2[22] + new_values3[22] + new_values4[22] + new_values5[22] +
                    + new_values6[22] + new_values7[22] + new_values8[22] + new_values9[22],
                    new_values1[23] + new_values2[23] + new_values3[23] + new_values4[23] + new_values5[23] +
                    + new_values6[23] + new_values7[23] + new_values8[23] + new_values9[23],
                    new_values1[24] + new_values2[24] + new_values3[24] + new_values4[24] + new_values5[24] +
                    + new_values6[24] + new_values7[24] + new_values8[24] + new_values9[24],
                    new_values1[25] + new_values2[25] + new_values3[25] + new_values4[25] + new_values5[25] +
                    + new_values6[25] + new_values7[25] + new_values8[25] + new_values9[25]]

    return new_values1, new_values2, new_values3, new_values4, new_values5, new_values6, new_values7, new_values8, new_values9, new_values10




ROOT_DIR = os.path.dirname(os.path.abspath(__file__)) + '\Данные'
file = filecsvlist(os.listdir(ROOT_DIR))
if len(file) > 1:
    input(f'В папке {ROOT_DIR} находится более одного файла. Удалите файлы и перезапустите программу'
          f'Для выхода нажмите Enter')
    exit()
elif len(file) < 1:
    input(f'В папке {ROOT_DIR} отсутствует файл для анализа. Перезапустите выполнение. Для выхода нажмите Enter')
    exit()
filedir = ROOT_DIR +'\\' + file[0]
print(filedir)
df = pd.read_excel(filedir)
df = df.loc[((df['1'] == 'ГУЗ') & (df['6'] == 'КРУГЛОСУТОЧНЫЙ СТАЦИОНАР') & (df['6'] == 'КРУГЛОСУТОЧНЫЙ СТАЦИОНАР') & (df['8'] == 'СПЕЦИАЛИЗИРОВАННАЯ') & (df['10'] == 'педиатрии'))
            | ((df['1'] == 'ГУЗ') & (df['6'] == 'КРУГЛОСУТОЧНЫЙ СТАЦИОНАР') & (df['6'] == 'КРУГЛОСУТОЧНЫЙ СТАЦИОНАР') & (df['8'] == 'СПЕЦИАЛИЗИРОВАННАЯ ВЫСОКОТЕХНОЛОГИЧНАЯ') & (df['10'] == 'педиатрии'))]
medorglist = df['4'].tolist()
medorglist = list(set(medorglist))
print(medorglist)
svodtabl = pd.DataFrame(columns=['Наименование медицинская организация', 'Профиль койки', 'Количество коек',
                                 'По МО без МТР Случаи', 'По МО без МТР Дни. посещ.', 'По МО без МТР Стоимость(руб)',
                                 'По межтерр помощи МТР Случаи', 'По межтерр помощи МТР Дни. посещ.',
                                 'По межтерр помощи МТР Стоимость(руб)', 'Всего Случаи', 'Всего Дни. посещ.',
                                 'Всего Стоимость(руб)', 'Нормативный показатель работы койки', 'Работа койки',
                                 'Расчетное значение коек', 'Профиль помощи', '!По МО без МТР Случаи',
                                 '!По МО без МТР Дни. посещ.', '!По МО без МТР Стоимость(руб)',
                                 '!По межтерр помощи МТР Случаи', '!По межтерр помощи МТР Дни. посещ.',
                                 '!По межтерр помощи МТР Стоимость(руб)', '!Всего Случаи', '!Всего Дни. посещ.',
                                 '!Всего Стоимость(руб)', '!Расчетное значение коек'])



for i in range(len(medorglist)):
    nv1, nv2, nv3, nv4, nv5, nv6, nv7, nv8, nv9, nv10 = org_tabl(medorglist[i])
    svodtabl.loc[len(svodtabl.index)] = nv1
    svodtabl.loc[len(svodtabl.index)] = nv2
    svodtabl.loc[len(svodtabl.index)] = nv3
    svodtabl.loc[len(svodtabl.index)] = nv4
    svodtabl.loc[len(svodtabl.index)] = nv5
    svodtabl.loc[len(svodtabl.index)] = nv6
    svodtabl.loc[len(svodtabl.index)] = nv7
    svodtabl.loc[len(svodtabl.index)] = nv8
    svodtabl.loc[len(svodtabl.index)] = nv9
    svodtabl.loc[len(svodtabl.index)] = nv10
    print(f'{i+1} из {len(medorglist)+1} выгружено')

svodtabl.to_excel('Анализ МО сводная.xlsx', index=False)