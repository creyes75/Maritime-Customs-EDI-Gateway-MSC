/* ejecutrar esto des la base que esta mas actualizada */
insert into docum..docu_buque
select * from docu_buque where buqu_linea = 'CAE' and buqu_codigo not in (select buqu_codigo from docum..docu_buque where buqu_linea ='CAE')
go
select count(*) from  docum..docu_buque
go
select count(*) from  docu_buque
go


insert into docum..docu_buque_viaje
select * from docu_buque_viaje a where a.buqv_linea = 'CAE' and a.buqv_viaje not in (select buqv_viaje from docum..docu_buque_viaje where buqv_linea ='CAE' and buqv_buque = a.buqv_buque)
go
select count(*) from  docum..docu_buque_viaje
go
select count(*) from  docu_buque_viaje
go


select count(*) from  docum..docu_blcab_exp
go
select count(*) from  docu_blcab_exp


