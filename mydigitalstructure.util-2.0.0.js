var _ = require('lodash')
var moment = require('moment');
var formdata = require('form-data');

module.exports = 
{
	VERSION: '1.0.0',

	data: {},

    attachment:
    {
        upload: function (param, fileData)
        {
            var filename = mydigitalstructure._util.param.get(param, 'filename', {default: 'export.xlsx'}).value;
            var object = mydigitalstructure._util.param.get(param, 'object', {default: 107}).value;
            var objectContext = mydigitalstructure._util.param.get(param, 'objectContext').value;
            var base64 = mydigitalstructure._util.param.get(param, 'base64', {default: false}).value;
            var type = mydigitalstructure._util.param.get(param, 'type').value;

            if (base64)
            {
                mydigitalstructure.cloud.invoke(
                {
                    method: 'core_attachment_from_base64',
                    data:
                    {
                        base64: fileData.base64,
                        filename: filename,
                        object: object,
                        objectcontext: objectContext
                    },
                    callback: module.exports.attachment.process,
                    callbackParam: param
                });
            }
            else
            {
                var settings = mydigitalstructure.get({scope: '_settings'});
                var session = mydigitalstructure.data.session;

                var blob = Buffer.from(fileData.buffer)

                var form = new formdata();
    
                form.append('file0', blob,
                {
                    contentType: 'application/octet-stream',
                    filename: filename
                });
                form.append('filename0', filename);
                form.append('object', object);
                form.append('objectcontext', objectContext);
                form.append('method', 'ATTACH_FILE');
                form.append('sid', session.sid);
                form.append('logonkey', session.logonkey);

                if (!_.isUndefined(type))
                {
                    form.append('type0', type);
                }

                var url = 'https://' + settings.mydigitalstructure.hostname + '/rpc/attach/'

                form.submit(url, function(err, res)
                {
                    res.resume();

                    res.setEncoding('utf8');
                    res.on('data', function (chunk)
                    {
                        var data = JSON.parse(chunk);
                        module.exports.attachment.process(param, data)
                    });
                });
            }
        },

        process: function (param, response)
        {
            if (response.status == 'OK')
            {
                var attachment;

                if (_.has(response, 'data.rows'))
                {
                    attachment = _.first(response.data.rows);
                }
                else
                {
                    attachment = response;
                }

                var data =
                {
                    attachment:
                    {
                        id: attachment.attachment,
                        link: attachment.attachmentlink,
                        href: '/download/' + attachment.attachmentlink
                    }
                }
            }

            param = mydigitalstructure._util.param.set(param, 'data', data);

            mydigitalstructure._util.complete(param)
        }
    }
}