"""Training and inference demand equations used by the study."""


def calc_Rt(nt_para, nt_traindata, nt_model, sp_rate, day):
    return nt_para * nt_traindata * nt_model / 86.4 / day / sp_rate


def calc_Ri(nt_para, ni_user, ni_query_amount, sp_rate):
    return nt_para * ni_user * ni_query_amount / 86.4 * 1.25 / sp_rate
