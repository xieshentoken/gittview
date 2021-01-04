# 输入十六进制颜色值string，输出一个包含(h,s,l)的list
def toHSL(rgb):
    if '#' in rgb: rgb = rgb[1:]
    r = int(rgb[:2], 16)/255
    g = int(rgb[2:4], 16)/255
    b = int(rgb[4:], 16)/255
    maxcolor = max(r, g, b)
    mincolor = min(r, g, b)

    l = (maxcolor+mincolor)/2
    
    if (l == 0)or(maxcolor == mincolor):
        s = 0
    elif (l > 0)and(l <= 0.5):
        s = (maxcolor-mincolor)/(maxcolor+mincolor)
    elif l > 0.5:
        s = (maxcolor-mincolor)/(2-maxcolor-mincolor)

    if maxcolor == mincolor:
        h = 0
    elif (maxcolor == r)and(g >= b):
        h = 60*(g-b)/(maxcolor-mincolor)
    elif (maxcolor == r)and(g < b):
        h = 60*(g-b)/(maxcolor-mincolor)+360
    elif maxcolor == g:
        h = 60*(b-r)/(maxcolor-mincolor)+120
    elif maxcolor == b:
        h = 60*(r-g)/(maxcolor-mincolor)+240
    
    return [round(h,3), round(s,3), round(l,3)]
    
# 输入一个[h,s,l]的list其值为float，输出十六进制的rgb值
def toRGB(hsl):
    h = hsl[0]/360
    s = hsl[1]
    l = hsl[2]
    r = g = b = 0
    rgbl = []
    if s == 0:
        r = g = b = round(l*255)
        rgb = '#' + str(hex(r))[2:]+ str(hex(g))[2:]+ str(hex(b))[2:]
        return rgb
    elif l < 0.5:
        q = l*(1+s)
    elif l >= 0.5:
        q = l+s-l*s
    p = 2*l-q
    t_r = h+1/3
    t_g = h
    t_b = h-1/3
    tt = []
    for t in [t_r, t_g, t_b]:
        if t < 0:
            t = t+1
        elif t > 1:
            t = t-1
        tt.append(t)
    for t, c in zip(tt,[r, g, b]):
        if t < 1/6:
            c = p+((q-p)*6*t)
        elif (t >= 1/6)and(t < 0.5):
            c = q
        elif (t >= 0.5)and(t < 2/3):
            c = p+((q-p)*6*(2/3-t))
        else:
            c = p
        rgbl.append(round(c*255))
    rgb = '#' + str(hex(rgbl[0]))[2:]+ str(hex(rgbl[1]))[2:]+ str(hex(rgbl[2]))[2:]
    return rgb