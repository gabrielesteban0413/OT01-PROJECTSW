_block
	_local ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_out.txt"
	_local ruta_salida << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_find.txt"

	_local input_file << external_text_input_stream.new(ruta_fuente)
	_local output_file << external_text_output_stream.new(ruta_salida)

	_local vista << gis_program_manager.cached_dataset(:gis)

	# Encabezado una sola vez
	output_file.write("SHEATH_ID|FUNDA|NOMBRE_ANTIGUO|HILO|ID_HILO|NOMBRE|ESTADO|TOPOLOGIA|TECNOLOGIA|STATUS|USO|CARACTERISTICA|RANGO|ANCHO_BANDA")
	output_file.newline()

	_for una_linea _over 1.upto(1000000)
	_loop
		_local linea << input_file.get_line()
		_if linea _is _unset
		_then
			_leave
		_endif

		_local id_texto << linea.write_string.trim_spaces()
		_if id_texto.size = 0
		_then
			_continue
		_endif

		_try
			_local id_sheath << id_texto.as_number()
			_local sheath << vista.collection(:sheath_with_loc).select(predicate.eq(:id, id_sheath)).an_element()

			_if sheath _is _unset
			_then
				output_file.write(id_texto + "|ERROR_NO_ENCONTRADA|-|-|-|-|-|-|-|-|-|-|-")
				output_file.newline()
				_continue
			_endif

			_local total_hilos << sheath.fiber_count
			_local nombre_funda << sheath.name.write_string
			_local line_counts << sheath.copper_line_of_counts

			# Nombre antiguo pertenece a la funda (sheath_with_loc), no al hilo
			_local nombre_antiguo_funda << sheath.nombre_antiguo
			_if nombre_antiguo_funda _is _unset
			_then
				nombre_antiguo_funda << "-"
			_else
				nombre_antiguo_funda << nombre_antiguo_funda.write_string
			_endif

			# Construir mapa fibra -> objeto
			_local fibra_a_objeto << rope.new()
			_for i _over 1.upto(total_hilos)
			_loop
				fibra_a_objeto.add(_unset)
			_endloop

			_if line_counts _isnt _unset
			_then
				_for elem _over line_counts.fast_elements()
				_loop
					_local low << elem.actual_low_range
					_local high << elem.actual_high_range
					_if low _isnt _unset _andif high _isnt _unset
					_then
						_for fibra _over low.upto(high)
						_loop
							_if fibra >= 1 _andif fibra <= total_hilos
							_then
								_if fibra_a_objeto[fibra] _is _unset
								_then
									fibra_a_objeto[fibra] << elem
								_endif
							_endif
						_endloop
					_endif
				_endloop
			_endif

			# Escribir una fila por hilo
			_for i _over 1.upto(total_hilos)
			_loop
				_local elem << fibra_a_objeto[i]
				_local nombre << ""
				_local estado << ""
				_local Id << "-"
				_local topologia << "-"
				_local tecnologia << "-"
				_local status << "-"
				_local uso << "-"
				_local caracteristica << "-"
				_local rango << "-"
				_local ancho << 0

				_if elem _isnt _unset
				_then
					Id << elem.id.write_string
					_if Id _is _unset _then Id << "-" _endif

					nombre << elem.designation
					_if nombre _is _unset _then nombre << "" _endif
					_if nombre = "UNDESIGNATED" _then nombre << "" _endif

					_if nombre = ""
					_then
						estado << "LIBRE"
					_else
						estado << "FUNCIONANDO"
					_endif

					topologia << elem.tipo_topologia
					_if topologia _is _unset _then topologia << "-" _endif
					tecnologia << elem.tipo_tecnologia
					_if tecnologia _is _unset _then tecnologia << "-" _endif
					status << elem.physical_status
					_if status _is _unset _then status << "-" _endif
					uso << elem.tipo_uso
					_if uso _is _unset _then uso << "-" _endif
					caracteristica << elem.characteristic
					_if caracteristica _is _unset _then caracteristica << "-" _endif

					_local low2 << elem.actual_low_range
					_local high2 << elem.actual_high_range
					_if low2 _isnt _unset _andif high2 _isnt _unset
					_then
						rango << low2.write_string + "-" + high2.write_string
					_endif

					ancho << elem.ancho_de_banda
					_if ancho _is _unset _then ancho << 0 _endif
				_else
					estado << "LIBRE"
				_endif

				_local fila << id_texto + "|" + nombre_funda + "|" + nombre_antiguo_funda + "|" +
				              i.write_string + "|" + Id + "|" + nombre + "|" + estado + "|" +
				              topologia + "|" + tecnologia + "|" + status + "|" +
				              uso + "|" + caracteristica + "|" + rango + "|" +
				              ancho.write_string

				output_file.write(fila)
				output_file.newline()
			_endloop

		_when error
			output_file.write(id_texto + "|ERROR_PROCESANDO|-|-|-|-|-|-|-|-|-|-|-")
			output_file.newline()
		_endtry
	_endloop

	input_file.close()
	output_file.close()

	show("--------TRACE_TABLE_MASIVO_LISTO---------")
_endblock