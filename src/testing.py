from usefulObjets import sapInterfaceJob

preApprovedParametersList = []
parametersList2 = []

l = ['z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z', 'z']
l0 = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'b', 'j', 'k', 'l', 'm', 'c', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w']
l1 = ['23.12.2022', '22.12.2022', '21.12.2022', '20.12.2022', '19.12.2022', '18.12.2022', '17.12.2022', '16.12.2022', '15.12.2022', '14.12.2022', '13.12.2022', '12.12.2022', '11.12.2022', '10.12.2022', '09.12.2022', '08.12.2022', '07.12.2022', '06.12.2022', '05.12.2022', '04.12.2022', '03.12.2022', '02.12.2022', '01.12.2022']
l2 = ['1', '2', '3', '4-', '5', '6-', '7', '8', '9-', '10', '11', '12', '13', '14-', '15', '16', '17', '18', '19', '20', '21-', '22', '23']
l5 = ['11111111111 BELZA GUTIERREZ CONDORI', '11111111111 JOSE LUIS ALEJANDRO RODRIGUEZ', '11111111111 NELVI JUANITA ROMERO', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 BELZA GUTIERREZ CONDORI', '11111111111 BELZA GUTIERREZ CONDORI', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 JOSE LUIS ALEJANDRO RODRIGUEZ', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 holis', '11111111111 holis']
print(len(l1))
print(l5[8])
print(len(l2))
print(len(l5))

parametersList2.append(l0)
parametersList2.append(l)
parametersList2.append(l1)
parametersList2.append(l)
parametersList2.append(l2)
parametersList2.append(l5)
parametersList2.append(l)
parametersList2.append(l)

l0_2 = ['a', 'b', 'c', 'd']
l_2 = ['z', 'z', 'z', 'z']
l3 = ['23.12.2022', '13.12.2022', '08.12.2022', '01.12.2022']
l4 = ['1-', '9', '14', '23-']

preApprovedParametersList.append(l0_2)
preApprovedParametersList.append(l_2)
preApprovedParametersList.append(l3)
preApprovedParametersList.append(l_2)
preApprovedParametersList.append(l4)
preApprovedParametersList.append(l_2)
preApprovedParametersList.append(l_2)
preApprovedParametersList.append(l_2)

# print(parametersList2)
# print(preApprovedParametersList)

x = sapInterfaceJob()
x.chargeListOfNames()
y = x.lastValidationChecker2(preApprovedParametersList, parametersList2)
print(y)