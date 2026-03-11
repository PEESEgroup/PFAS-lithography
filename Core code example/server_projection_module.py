import numpy as np
import pandas as pd
import xlrd
from itertools import starmap as smap


def cum_p_outflow(time_index, P, S_in, len_lifespan, lifespan_distribution):
    cum_result = 0
    index = min(time_index, len_lifespan)
    for j in range(index):
        cum_result = cum_result + P[time_index - j - 1] * S_in[time_index - j - 1] * float(lifespan_distribution[j])
    return cum_result


def cum_s_outflow(time_index, S_in, len_lifespan, lifespan_distribution):
    cum_result = 0
    index = min(time_index, len_lifespan)
    for j in range(index):
        cum_result = cum_result + S_in[time_index - j - 1] * float(lifespan_distribution[j])
    return cum_result


def fcn1(i, P):
    if P[i] == P[i - 1]:
        return 1
    return 0


def randomize(n, sigma):
    return np.random.normal(n, sigma)


def calc_sigma_002(n):
    return n * 0.01


def calc_sigma_020(n):
    return n * 0.20


def calc_sigma_005(n):
    return n * 0.05


def calc_Rt(nt_para, nt_traindata, nt_model, sp_rate, day):
    return nt_para * nt_traindata * nt_model / 86.4 / day / sp_rate


def calc_Ri(nt_para, ni_user, ni_query_amount, sp_rate):
    return nt_para * ni_user * ni_query_amount / 86.4 * 1.25 / sp_rate


def run_server_projection(input_path='data input/data.xls'):
    S1_total = pd.DataFrame()
    S2_total = pd.DataFrame()
    C1_total = pd.DataFrame()
    C2_total = pd.DataFrame()
    R_total = pd.DataFrame()
    Out_total = pd.DataFrame()

    excel = xlrd.open_workbook(input_path)
    config_data = excel.sheet_by_index(0)
    sheet1 = excel.sheet_by_index(1)
    sheet2 = excel.sheet_by_index(2)

    Nt_para = sheet1.col_values(colx=0)
    Nt_para.pop(0)
    Nt_traindata = sheet1.col_values(colx=1)
    Nt_traindata.pop(0)
    Pt_theo = sheet1.col_values(colx=2)
    Pt_theo.pop(0)
    Pi_theo = sheet1.col_values(colx=3)
    Pi_theo.pop(0)
    Ni_user = sheet1.col_values(colx=4)
    Ni_user.pop(0)
    Eff_t = sheet1.col_values(colx=5)
    Eff_t.pop(0)
    Eff_i = sheet1.col_values(colx=6)
    Eff_i.pop(0)
    Sp_rate = sheet1.col_values(colx=7)
    Sp_rate.pop(0)
    N_model = sheet1.col_values(colx=8)
    N_model.pop(0)

    lifespan_distribution = sheet2.col_values(colx=0)
    lifespan_distribution.pop(0)

    random_batch = int(config_data.cell(0, 1).value)
    l_range = int(config_data.cell(1, 1).value)
    day = int(config_data.cell(2, 1).value)
    GPU_per_server = int(config_data.cell(3, 1).value)
    upgrade_strategy = int(config_data.cell(4, 1).value)
    n_token = int(config_data.cell(5, 1).value)

    lifespan_distribution = [0] * int(l_range - len(lifespan_distribution) / 2) + lifespan_distribution + [0] * int(l_range - len(lifespan_distribution) / 2)

    T = len(Pt_theo)
    len_lifespan = len(lifespan_distribution)
    l = l_range

    S1_runs = []
    S2_runs = []
    C1_runs = []
    C2_runs = []
    R_runs = []
    Out_runs = []

    for _ in range(random_batch):
        S1_t = np.zeros(T)
        S2_t = np.zeros(T)
        S1_i = np.zeros(T)
        S2_i = np.zeros(T)
        Out_t = np.zeros(T)
        Out_i = np.zeros(T)
        zzz_t = np.zeros(T)
        zzz_i = np.zeros(T)

        zzz_ii = np.zeros(T)
        eff_t = list(Eff_t)
        eff_i = list(Eff_i)
        sp_rate = list(Sp_rate)

        ni_query_amount = randomize(n_token, 100)
        nt_para = list(smap(randomize, zip(Nt_para, list(map(calc_sigma_002, Nt_para)))))
        nt_traindata = list(smap(randomize, zip(Nt_traindata, list(map(calc_sigma_002, Nt_traindata)))))
        ni_user = list(smap(randomize, zip(Ni_user, list(map(calc_sigma_002, Ni_user)))))
        pt_theo = list(smap(randomize, zip(Pt_theo, list(map(calc_sigma_002, Pt_theo)))))
        pi_theo = list(smap(randomize, zip(Pi_theo, list(map(calc_sigma_002, Pi_theo)))))

        R_t = list(smap(calc_Rt, zip(nt_para, nt_traindata, N_model, sp_rate, [day] * T)))
        R_i = list(smap(calc_Ri, zip(nt_para, ni_user, [ni_query_amount] * T, sp_rate)))

        R_t.insert(0, 0)
        R_i.insert(0, 0)
        P_t = [x * eff_t[i] * GPU_per_server for i, x in enumerate(pt_theo)]
        P_i = [x * eff_i[i] * GPU_per_server for i, x in enumerate(pi_theo)]

        S1_t[0] = np.ceil((R_t[1] - R_t[0]) / P_t[0])
        S1_i[0] = np.ceil((R_i[1] - R_i[0]) / P_i[0])
        for i in range(1, T):
            Out_t[i] = cum_s_outflow(i, S1_t, len_lifespan, lifespan_distribution)
            S1_t[i] = np.ceil((R_t[i + 1] - R_t[i] + cum_p_outflow(i, P_t, S1_t, len_lifespan, lifespan_distribution)) / P_t[i])
            zzz_t[i] = cum_p_outflow(i, P_t, S1_t, len_lifespan, lifespan_distribution) / P_t[i]

            S2_t[i] = np.ceil((R_t[i + 1] - fcn1(i, P_t) * R_t[i]) / P_t[i])

            Out_i[i] = cum_s_outflow(i, S1_i, len_lifespan, lifespan_distribution)
            S1_i[i] = np.ceil((R_i[i + 1] - R_i[i] + cum_p_outflow(i, P_i, S1_i, len_lifespan, lifespan_distribution)) / P_i[i])
            zzz_i[i] = cum_p_outflow(i, P_i, S1_i, len_lifespan, lifespan_distribution) / P_i[i]
            zzz_ii[i] = R_i[i + 1] - R_i[i]
            S2_i[i] = np.ceil((R_i[i + 1] - fcn1(i, P_i) * R_i[i]) / P_i[i])

        S1 = [S1_t[i] + S1_i[i] for i in range(T)]
        S2 = [S2_t[i] + S2_i[i] for i in range(T)]
        Out = [Out_t[i] + Out_i[i] for i in range(T)]

        C1 = np.zeros(T)
        C2 = np.zeros(T)
        C1[0] = S1[0]
        for i in range(T):
            if i < l:
                C1[i] = C1[i - 1] + S1[i]
            else:
                C1[i] = C1[i - 1] + S1[i] - S1[i - l]
            C2[i] = np.ceil(R_t[i + 1] / P_t[i] + R_i[i + 1] / P_i[i])

        S1_runs.append(S1)
        S2_runs.append(S2)
        C1_runs.append(C1)
        C2_runs.append(C2)
        R_runs.append(list(np.add(R_t, R_i)))
        Out_runs.append(Out)

    cols = list(range(random_batch))

    S1_total = pd.DataFrame(S1_runs).T
    S1_total.columns = cols

    S2_total = pd.DataFrame(S2_runs).T
    S2_total.columns = cols

    C1_total = pd.DataFrame(C1_runs).T
    C1_total.columns = cols

    C2_total = pd.DataFrame(C2_runs).T
    C2_total.columns = cols

    R_total = pd.DataFrame(R_runs).T
    R_total.columns = cols

    Out_total = pd.DataFrame(Out_runs).T
    Out_total.columns = cols

    if upgrade_strategy == 1:
        S_final = S1_total
    else:
        S_final = S2_total

    mean_list = []
    std_list = []
    std_cum_list = []
    mean_R = []
    mean_out = []
    std_out_list = []
    std_cum_out_list = []
    std_R = []
    for i in range(T):
        mean_list.append(np.mean(S_final.loc[i, :]))
        std_list.append(np.std(S_final.loc[i, :], ddof=1))
        std_cum_list.append(np.std([np.sum(S_final.loc[0:i, j]) for j in range(random_batch)]))
        mean_out.append(np.mean(Out_total.loc[i, :]))
        std_out_list.append(np.std(Out_total.loc[i, :], ddof=1))
        std_cum_out_list.append(np.std([np.sum(Out_total.loc[0:i, j]) for j in range(random_batch)]))
    for i in range(T + 1):
        mean_R.append(np.mean(R_total.loc[i, :]))
        std_R.append(np.std(R_total.loc[i, :], ddof=1))

    return {
        'mean_list': mean_list,
        'std_list': std_list,
        'std_cum_list': std_cum_list,
        'mean_out': mean_out,
        'std_out_list': std_out_list,
        'std_cum_out_list': std_cum_out_list,
        'mean_R': mean_R,
        'std_R': std_R,
    }
