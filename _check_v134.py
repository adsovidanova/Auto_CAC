with open(r'C:\Users\PO-TC1ED\OneDrive - 901388331_INST. DE SERVICIOS INTEGRALES EN CANCER\Desktop\Auto_CAC\20260505_ESS062_CANCER.txt','r',encoding='cp1252') as f:
    h=f.readline().split('\t')
    idx=h.index('VAR134_FechaCorte')
    for i,l in enumerate(f):
        if i<3:
            print(l.split('\t')[idx])
        else:
            break
