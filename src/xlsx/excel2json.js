/**
 * [description]
 * excel配置会议议程
 * @author mao
 * @version 1
 * @date    2016-09-28
 */
var excel = (function(mod) {

	/**
	 * [checkColumn description]
	 * 检测是几列数据
	 * @author mao
	 * @version 1
	 * @date    2016-10-12
	 * @param   {object}   workbook excel数据
	 * @return  {[object]}            结果数据
	 */
	function checkColumn(workbook) {
		var result = workbook.Sheets[(workbook.SheetNames)[0]],
				column,
				data;

		//判断是列excel数据
		column = ((result['!ref']).split(':')[1]).charAt(0);

		//处理成hash结构
		data = processOrigin(result, column);

		return data;
	}

	/**
	 * [processOrigin description]
	 * @author mao
	 * @version 1
	 * @date    2016-10-12
	 * @param   {object}   result 待处理excel数据
	 * @param   {string}   column 有多少列数据
	 * @return   {Object}    结果数据
	 */
	function processOrigin(result, column) {
		var merges = result['!merges'] || [],  //合并表格的位置信息
				obj,response;

		//生成对应的几项hash数据
		switch(column) {
			case 'B' : 
				obj = {
					_1:{}
				};
				break;
			case 'C' :
			  obj = {
					_1:{},
					_2:{}
				};
			  break;
			case 'D' :
			  obj = {
					_1:{},
					_2:{},
					_3:{}
				};
			  break;
			case 'E' :
			  obj = {
					_1:{},
					_2:{},
					_3:{},
					_4:{}
				};
			  break;
			default: break;
		}

		//拿到excel初始的hash数据
		for(var i in result) {
			if(i.charAt(0) === '!' || i.charAt(0) === 'A') continue;
			switch(i.charAt(0)) {
				case 'B':{
					var key = i.slice(1,i.length);
					obj._1[key] = result[i].v;
					break;
				}
				case 'C':{
					var key = i.slice(1,i.length);
					obj._2[key] = result[i].v;
					break;
				}
				case 'D':{
					var key = i.slice(1,i.length);
					obj._3[key] = result[i].v;
					break;
				}
				case 'E':{
					var key = i.slice(1,i.length);
					obj._4[key] = result[i].v;
					break;
				}
				default:break;
			}
		}

		//合并项
		response = mergeColumn(obj, merges);

		return response;
	}

	/**
	 * [mergeColumn]
	 * description
	 * @author mao
	 * @version 1
	 * @date    2016-10-12
	 * @param   {obj}   obj    取出的excel数据
	 * @param   {Object}   merges 合并单元格的起始坐标
	 * @return  {Object}          补全后的单元格数据
	 */
	function mergeColumn(obj, merges) {
		//判断是否只为一列
		var _keys = [];
		for(var i in obj) {
			_keys.push(i);
		}
		if(_keys.length === 1) {
			return obj;
		}

		//验证是否有合并
		if(merges.length === 0) {
			if(typeof console !== 'undefined') {
				console.log('merges is empty');
			}
			return obj;
		}

		//将数据处理成全项目的hash
		for(var i = 0; i < merges.length; i++) { //纵向合并
			if(merges[i].e.c == merges[i].s.c) { 
				var start = merges[i].s.r + 1,
						end = merges[i].e.r + 1,
						sub = merges[i].e.c, //起点x坐标
						range = end - start,
						origin = obj['_' + sub][start];
				
				//起始点数据
				obj['_' + sub][start] = {
					_v: origin,
					_w: 'row',
					_s: true,
					_c: (range + 1)
				}

				//补全被合并项
				for(var j = 1; j <= range; j++) {
					start ++;
					obj['_' + sub][start] = {
						_v: origin,
						_w: 'row',
					  _s: false,
						_c: (range + 1)
					}
				}

			} else { //横向的合并
				var start = merges[i].s.c,
						end = merges[i].e.c,
						sub = merges[i].e.r + 1, //起点y坐标
						range = end - start,
						origin = obj['_' + start][sub];

				//起始点数据
				obj['_' + start][sub] = {
					_v: origin,
					_w: 'col',
					_s: true,
					_c: (range + 1)
				}

				//补全被合并项
				for(var j = 1; j <= range; j++) {
					start ++;
					obj['_' + start][sub] = {
						_v: origin,
						_w: 'col',
						_s: false,
						_c: (range + 1)
					}
				}
			}
		}

		return obj;
	}

	/**
	 * 数组排序
	 * @author mao
	 * @version 1.01
	 * @date    2016-11-03
	 * @param   {number}   a 排序依据
	 * @param   {number}   b 排序依据
	 * @return  {array}     排序后的数组
	 */
	function sortKEY(a,b) {
		return a._id - b._id;
	}

	/**
	 * hash数据整理 为兼容IE8及以下浏览器
	 * @author mao
	 * @version 1.01
	 * @date    2016-11-03
	 * @param   {Object}   data 待整理的hash
	 * @return  {Object}        整理后的hash
	 */
	function sortHash(data) {
		var arr = [],
				obj = {},
				result;

		for(var i in data) {
			arr.push({_id:i,_value:data[i]});
		}

		result = arr.sort(sortKEY);
		for(var i = 0; i < result.length; i++) {
			obj[result[i]._id] = result[i]._value;
		}

		return obj;
	}

	/**
	 * [toArray description]
	 * @author mao
	 * @version 1
	 * @date    2016-10-12
	 * @param   {Object}   obj 待处理的hash
	 * @return  {array}       处理成的数组
	 */
	function toArray(obj) {
		var keys = [],
				data = sortHash(obj._1),  //为兼容IE8及以下
				res = [];
		//获取key值
		for(var i in obj) {
			keys.push(i);
		}

		//处理成数组
		for(var i in data) {
			var current = {};
			for(var j = 0; j < keys.length; j++) {
				current[keys[j]] = obj[keys[j]][i];
			}
			res.push(current);
		}
		
		return res;
	}


	/**
	 * [createXHR]
	 * 创建一个xhr
	 * @author mao
	 * @version 1
	 * @date    2016-09-26
	 * @return  {object}   xhr
	 */
	function createXHR() {
		if(window.XMLHttpRequest) {
	    return new XMLHttpRequest();
		} else if(window.ActiveXObject) {  //ie6
			return new ActiveXObject('MSXML2.XMLHTTP.3.0');
		} else {
			throw 'XHR unavailable for your browser';
		}
	}


	/**
	 * [transferData]
	 * 请求excel文件请求
	 * @author mao
	 * @version 1
	 * @date    2016-09-28
	 * @param   {Function} option.dataRender 回调函数，处理结果数据
	 * @param   {string}   option.url      xlsx文件请求地址
	 */
	mod.transferData = function(option) {
		//新建xhr
		var oReq = createXHR();
		//建立连接
		oReq.open("GET", option.url, true);

		//判断是否为低版本的ie，处理返回
		if(typeof Uint8Array !== 'undefined') {
			oReq.responseType = "arraybuffer";
			oReq.onload = function(e) {
				if(typeof console !== 'undefined') console.log("onload", new Date());
				var arraybuffer = oReq.response;
				var data = new Uint8Array(arraybuffer);
				var arr = new Array();
				for(var i = 0; i != data.length; ++i) {
					arr[i] = String.fromCharCode(data[i]);
				}
				//处理数据
				var wb = XLSX.read(arr.join(""), {type:"binary"});

				//to_json
				if(typeof option.toJson == 'function') {
					option.toJson(XLSX.utils.sheet_to_row_object_array(wb.Sheets[(wb.SheetNames)[0]]));
				}
				//数据放入回调
				if(typeof option.dataRender == 'function') {
					option.dataRender(toArray(checkColumn(wb)));
				}
			};
		} else {
			oReq.setRequestHeader("Accept-Charset", "x-user-defined");	
			oReq.onreadystatechange = function() { 
				if(oReq.readyState == 4 && oReq.status == 200) {
					var ff = convertResponseBodyToText(oReq.responseBody);
					if(typeof console !== 'undefined') {
						console.log("onload", new Date());
					}

					//处理数据
					var wb = XLSX.read(ff, {type:"binary"});

					//to_json
					if(typeof option.toJson == 'function') {
						option.toJson(XLSX.utils.sheet_to_row_object_array(wb.Sheets[(wb.SheetNames)[0]]));
					}
					//数据放入回调
					if(typeof option.dataRender == 'function') {
						option.dataRender(toArray(checkColumn(wb)));
					}
				}
			};
		}

		//发送请求
		oReq.send();
	}

	/**
	 * [check_undefind description]
	 * @author mao
	 * @version 1
	 * @date    2016-10-13
	 * @param   {string}   data 数据
	 * @return  {string}        返回空或数据本身
	 */
	mod.check_undefind = function(data) {
		if(!data) {
			return '';
		} else {
			if(typeof data != 'number') {
				//检测特殊字符
				if(data.indexOf('&#10;') != -1) {
					var results = data;
					if(results.indexOf(' ') != -1) {
						results = results.split(' ').join('<span class="blank"></span>');
					}
					return results.split('&#10;').join('<br/>');
				} else {
					return data.split(' ').join('&nbsp;');
				}
				return data;

			} else {
				return data;
			}
		}
	}

	/**
	 * [renderHTML description]
	 * @author mao
	 * @version 1
	 * @date    2016-10-13
	 * @param   {object}   table 最终数据
	 * @return  {string}         渲染的dom
	 */
	mod.renderHTML = function(table) {
		var html = '';
		for(var i = 0; i <table.length; i++) {
			html += '<tr>';
			for(var j in table[i]) {
				var item = table[i][j];
				if(typeof item === 'object') {
					switch(item._w) {
						case 'col': {
							if(item._s) {
								html += '<td colspan="'+item._c+'">'+mod.check_undefind(item._v)+'</td>';
							}
							break;
						}
						case 'row': {
							if(item._s) {
								html += '<td rowspan="'+item._c+'">'+mod.check_undefind(item._v)+'</td>';
							}
							break;
						}
						default:break;
					}
				} else {
					html += '<td>'+mod.check_undefind(item)+'</td>';
				}
			}
			html += '</tr>';
		}
		return html;
	}

	return mod;

})(excel || {})