var gulp = require('gulp'); 
var concat = require('gulp-concat');
var uglify = require('gulp-uglify');
var rename = require('gulp-rename');

// 合并，压缩文件
gulp.task('scripts', function() {
    gulp.src('./src/xlsx/*.js')
        .pipe(concat('xls2json.js'))
        .pipe(gulp.dest('./dist'))
        .pipe(rename('xls2json.min.js'))
        .pipe(uglify())
        .pipe(gulp.dest('./dist'));
});


gulp.task('default', function(){
    gulp.run('scripts');
});