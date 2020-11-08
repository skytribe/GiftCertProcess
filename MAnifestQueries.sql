


select  sCust, nBodyCnt, sItem, sComment ,  sFirstName, sLastName ,*
                     From tInv Inner Join 
                     tPeople ON tInv.wCustid = tPeople.wCustId 
                      Inner Join tPeopleAncillary  ON tInv.wCustid = tPeopleAncillary .wCustId 
                     Inner Join tPrices on tInv.wItemId = tPrices.wItemId 
                     --Where nMani = " & CStr(muaMani(i).lngManiId) & _
                  --WHERE nBodyCnt > 0 AND sItem like 'TJ%'
ORDER BY sCust



select   sCust, 
nBodyCnt, 
sItem, 
sComment ,
sFirstName,
sLastname,
*
                     From tInvaLL 
                     Inner Join tPeople ON tInvAll.wCustid = tPeople.wCustId  
           
                     Inner Join tPrices on tInvAll.wItemId = tPrices.wItemId 
                      Inner Join tPeopleAncillary  ON tInvAll.wCustid = tPeopleAncillary .wCustId 
                      --Where nMani = " & CStr(muaMani(i).lngManiId) & _
                      where dtprocess='2020-02-29 00:00:00.000' AND
                      nBodyCnt > 0 AND sItem like 'TJ%'
ORDER BY   nMani, sCust


Select tplane.sName, nriders, dtDepart, * from tManiAll
inner join tPlane on nPlaneId = tplane.nId 
  where dtprocess='2020-02-29 00:00:00.000'  