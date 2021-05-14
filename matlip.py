import matplotlib.pyplot as plt
country = [ 'egypt', 'ksa', 'yemen', 'dammam','kfs','usa','suda']
gdp = [-500,-900,-1000, -2000,-2500,-3000,-4000]
new_col = ['green', 'blue', 'purple','yellow']
plt.bar(country, gdp,color=new_col)
plt.title('GDP for arab country from 2000 to 2021')
plt.xlabel('country')
plt.ylabel('gdp')
plt.grid(False)
plt.show()