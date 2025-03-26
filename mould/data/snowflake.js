function flake() {
    //先写静态再写动态
    //一朵雪花在屏幕内随即摆放
    var f = document.createElement("img");
    f.src = "data//snowflake.png";
    //随机问题用随机函数
    //先获取屏幕的宽高，在用随机函数得到一个随机的X Y值
    var width = document.documentElement.clientWidth;
    var heigh = document.documentElement.clientHeight+3200;
    //获取屏幕随机坐标
    var left = Math.random() * width;
    var top = Math.random() * heigh;
    //  alert(width)
    f.style.position = "absolute";
    f.style.left = left + "px";
    f.style.top = top + "px";
    //随即缩放
    f.style.transform = "scale(" + Math.random() / 8 + ")"
    //将这个标签插入到body中
    document.body.appendChild(f);
    //在JS中可以使用方法里面的方法
    function down() {
        top++;
        left++;
        f.style.left = left + "px";
        f.style.top = top + "px";
        if (top > heigh) {
            top = -100;
        }
        if (left > width) {
            left = -100;
        }
    }
    setInterval(down, 20)
}
//下落
function down() {
    f.style.top++
}
setInterval(down, 1000)
for (var i = 0; i < 50; i++) {
    flake()
}
