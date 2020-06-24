-- spRelAuditoria '002','055','2004'
  
ALTER procedure dbo.spRelAuditoria(
@sig_orgao_processo char(3), 
@seq_processo varchar(3), 
@ano_processo varchar(4))
as                
                
set nocount on                
          
declare @seq_processoaux tinyint, @ano_processoaux smallint          
declare @flag char(1)    
          
set @seq_processoaux = convert(tinyint, @seq_processo)          
set @ano_processoaux = convert(smallint, @ano_processo)          
                
--**********************************************************************                  
-- para pegar Processo AUDIN, data da SA, encaminhamento, Base Legal e euipe auditora para                
-- a geração da SA (NIG-003-Rev 00 Apr. Jan/03)                
--**********************************************************************                  
        
select 'PA' + '-' + p.sig_orgao_processo + '-' + right('000' + convert(varchar(3),p.seq_processo),3) + '/' + convert(varchar(4),p.ano_processo) + '-' + p.sig_tipo_auditoria as ProcessoAudin,                  
p.dta_inicio as  datainicio, p.dta_fim datafim , p.num_oficio numoficio, p.dta_oficio dataoficio,            
vuo.sig_uo as orgaoauditado, vuo.nom_uo + ' - ' + vuo.sig_uo orgao_auditado,
space(1000) as encaminhamento,  descr_auditor_responsavel as nomechefe, descr_conselho_auditor_responsavel as crcchefe, descr_func_auditor_responsavel funcchefe,      
nom_pessoa equipeauditora,vp.sig_uo_lotacao, '1' as flag into #Temp        
from processo_auditoria p               
 inner join vw_uo_auditoria vuo                  
 on p.cod_uo = vuo.cod_uo                  
 inner join equipe_auditora e                  
 on p.sig_orgao_processo =  e.sig_orgao_processo                   
 and p.seq_processo =  e.seq_processo                  
 and p.ano_processo = e.ano_processo                  
 inner join vw_pessoa_auditoria vp                
 on e.cod_pessoa_inmetro = vp.cod_pessoa_inmetro                  
 inner join vw_uo_auditoria vw  
 on vw.sig_uo = vp.sig_uo_lotacao  
where                   
 p.sig_orgao_processo =  @sig_orgao_processo                   
 and p.seq_processo =  @seq_processoaux                  
 and p.ano_processo = @ano_processoaux                  
 and p.ind_desativacao = 'A'                
        
declare @campo varchar(50)        
declare @encaminhamento varchar(1000)        
set @encaminhamento = ''        
        
DECLARE areaencaminha CURSOR FOR         
SELECT  case when not dbo.encaminhamento_relatorio.cod_uo is null then dbo.vw_uo_auditoria.sig_uo          
   else dbo.area_encaminha_externo.descr_area_encaminha_externo end as encaminhamento          
FROM         dbo.encaminhamento_relatorio left JOIN          
                      dbo.area_encaminha_externo ON           
                      dbo.encaminhamento_relatorio.seq_area_encaminha_externo = dbo.area_encaminha_externo.seq_area_encaminha_externo left JOIN          
                      dbo.vw_uo_auditoria          
        on dbo.encaminhamento_relatorio.cod_uo = dbo.vw_uo_auditoria.cod_uo          
WHERE     (dbo.encaminhamento_relatorio.ano_processo = @ano_processoAux) AND (dbo.encaminhamento_relatorio.seq_processo = @seq_processoAux) AND           
                      (dbo.encaminhamento_relatorio.sig_orgao_processo = @sig_orgao_processo)        
OPEN areaencaminha        
        
FETCH NEXT FROM areaencaminha         
INTO @campo        
WHILE @@FETCH_STATUS = 0        
 BEGIN        
  set @encaminhamento  = @encaminhamento + rtrim(@campo) + ';' + char(10)        
  FETCH NEXT FROM areaencaminha INTO @campo        
           
    END        
        
 CLOSE areaencaminha        
    DEALLOCATE areaencaminha        
        
    
if exists(select seq_comentario from comentario_item_sa 
 where sig_orgao_processo =  @sig_orgao_processo
 and seq_processo =  @seq_processoaux                  
 and ano_processo = @ano_processoaux
 and not descr_recomendacao is null
 and ind_desativacao = 'A'    
 and ind_relatorio = '1')    
set @flag='1'    
else    
set @flag='0'    
    
    
        
update #temp set encaminhamento = @encaminhamento, flag = @flag         
        
select * from #Temp        
drop table #Temp

