use docum
go

update docu_tipo_envio set tipo_env_descripcion = '01 - CANCELACION MANIFIESTO COMPLETO'
where tipo_env_linea = 'CAE' and tipo_env_codigo = '01'
go

update docu_tipo_envio set tipo_env_descripcion = '17 - CANCELACION BL'
where tipo_env_linea = 'CAE' and tipo_env_codigo = '17'
go

update docu_tipo_envio set tipo_env_descripcion = '20 - REEMPLAZO DATOS MANIFIESTO'
where tipo_env_linea = 'CAE' and tipo_env_codigo = '20'
go

update docu_tipo_envio set tipo_env_descripcion = '21 - REEMPLAZO DATOS BL'
where tipo_env_linea = 'CAE' and tipo_env_codigo = '21'
go

select * from docu_tipo_envio
go
